from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import pandas as pd


WORKSPACE_DIR = Path(r"C:\Users\User\Downloads\UNI\summer training\automation project 1")


def resolve_source_file() -> Path:
    preferred = WORKSPACE_DIR / "rebootPendingList (3) 1 with owners.xlsx"
    if preferred.exists():
        return preferred

    candidates = [
        path
        for path in WORKSPACE_DIR.glob("rebootPendingList (3) 1 with owners*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("Could not find the pending reboot workbook with owners.")

    return max(candidates, key=lambda path: path.stat().st_mtime)


def write_output(writer_factory, output_file: Path, summary: pd.DataFrame, owner_frames: dict[str, pd.DataFrame]) -> Path:
    try:
        with writer_factory(output_file) as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            used_names: set[str] = {"Summary"}
            for owner_name, owner_frame in owner_frames.items():
                sheet_name = sanitize_sheet_name(owner_name, used_names)
                owner_frame.to_excel(writer, index=False, sheet_name=sheet_name)
                used_names.add(sheet_name)
        return output_file
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_file = output_file.with_name(f"{output_file.stem}_{timestamp}{output_file.suffix}")
        with writer_factory(fallback_file) as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            used_names = {"Summary"}
            for owner_name, owner_frame in owner_frames.items():
                sheet_name = sanitize_sheet_name(owner_name, used_names)
                owner_frame.to_excel(writer, index=False, sheet_name=sheet_name)
                used_names.add(sheet_name)
        return fallback_file


def normalize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in dataframe.columns:
        normalized = str(column).replace("\ufeff", "").strip().strip('"')
        if normalized != column:
            rename_map[column] = normalized

    if rename_map:
        dataframe = dataframe.rename(columns=rename_map)

    return dataframe


def normalize_owner(value: object) -> str:
    if pd.isna(value):
        return "Unassigned"

    text = str(value).replace("\ufeff", "").strip().strip("'\"")
    if not text or text.lower() in {"--", "none", "nan", "null"}:
        return "Unassigned"

    return text


def sanitize_sheet_name(name: str, used_names: set[str]) -> str:
    base_name = re.sub(r"[\\/?*\[\]:]", "_", name).strip() or "Unassigned"
    base_name = base_name[:31]

    if base_name not in used_names:
        return base_name

    suffix = 1
    while True:
        trimmed_base = base_name[: 31 - len(f"_{suffix}")]
        candidate = f"{trimmed_base}_{suffix}"
        if candidate not in used_names:
            return candidate
        suffix += 1


def build_owner_frames(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    dataframe = normalize_columns(dataframe)

    if "Owner" not in dataframe.columns:
        raise KeyError("The source workbook does not contain an 'Owner' column.")

    preferred_order = [
        "Computer Name",
        "Domain",
        "IP Address",
        "Operating System",
        "Last Boot Time",
        "Owner",
        "Owner Email",
        "Notes",
        "Service Pack",
        "Health",
        "Agent Version",
        "Last Contact Time",
        "Reboot Required Reason",
    ]
    server_columns = [column for column in preferred_order if column in dataframe.columns]
    server_columns.extend(column for column in dataframe.columns if column not in server_columns)

    summary_rows: list[dict[str, object]] = []
    owner_frames: dict[str, pd.DataFrame] = {}

    for owner_name, owner_frame in dataframe.assign(_Owner=dataframe["Owner"].map(normalize_owner)).groupby("_Owner", sort=True):
        owner_frame = owner_frame.loc[:, server_columns].copy().reset_index(drop=True)
        owner_frames[owner_name] = owner_frame
        summary_rows.append({"Owner": owner_name, "Server Count": int(len(owner_frame))})

    summary = pd.DataFrame(summary_rows).sort_values(["Server Count", "Owner"], ascending=[False, True]).reset_index(drop=True)
    return summary, owner_frames


def print_owner_list(dataframe: pd.DataFrame) -> None:
    dataframe = normalize_columns(dataframe)

    owner_column = "Owner"
    computer_column = "Computer Name" if "Computer Name" in dataframe.columns else None
    ip_column = "IP Address" if "IP Address" in dataframe.columns else None
    email_column = "Owner Email" if "Owner Email" in dataframe.columns else None

    if owner_column not in dataframe.columns:
        raise KeyError("The source workbook does not contain an 'Owner' column.")
    if computer_column is None:
        raise KeyError("The source workbook does not contain a 'Computer Name' column.")

    grouped = dataframe.assign(_Owner=dataframe[owner_column].map(normalize_owner)).groupby("_Owner", sort=True)

    for owner_name, owner_frame in grouped:
        emails: list[str] = []
        if email_column is not None:
            emails = sorted({
                normalize_owner(value)
                for value in owner_frame[email_column]
                if normalize_owner(value) != "Unassigned"
            })

        email_text = ", ".join(emails) if emails else "n/a"
        servers = [
            (
                str(server_name).replace("\ufeff", "").strip().strip("'\""),
                str(ip_value).replace("\ufeff", "").strip().strip("'\""),
            )
            for server_name, ip_value in zip(
                owner_frame[computer_column],
                owner_frame[ip_column] if ip_column is not None else [""] * len(owner_frame),
            )
            if str(server_name).strip()
        ]

        print(f"Owner: {owner_name}")
        print(f"Email: {email_text}")
        for server_name, server_ip in servers:
            if server_ip and server_ip != "--":
                print(f"  - {server_name} ({server_ip})")
            else:
                print(f"  - {server_name}")
        print()


def main() -> None:
    source_file = resolve_source_file()
    dataframe = pd.read_excel(source_file)
    summary, owner_frames = build_owner_frames(dataframe)
    print_owner_list(dataframe)

    output_file = WORKSPACE_DIR / "rebootPendingList (3) 1 grouped by owner.xlsx"
    written_file = write_output(lambda path: pd.ExcelWriter(path, engine="openpyxl"), output_file, summary, owner_frames)

    if written_file == output_file:
        print(f"Wrote {written_file}")
    else:
        print(f"Target file was locked, wrote fallback file: {written_file}")


if __name__ == "__main__":
    main()