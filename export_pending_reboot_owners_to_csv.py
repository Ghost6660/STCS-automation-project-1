from __future__ import annotations

from pathlib import Path

import pandas as pd


WORKSPACE_DIR = Path(r"C:\Users\User\Downloads\UNI\summer training\automation project 1")
SOURCE_FILE = WORKSPACE_DIR / "rebootPendingList (3) 1 with owners.xlsx"
OUTPUT_FILE = WORKSPACE_DIR / "rebootPendingList (3) 1 with owners listed.csv"


def normalize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in dataframe.columns:
        normalized = str(column).replace("\ufeff", "").strip().strip('"')
        if normalized != column:
            rename_map[column] = normalized

    if rename_map:
        dataframe = dataframe.rename(columns=rename_map)

    return dataframe


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\ufeff", "").strip().strip("'\"")


def normalize_owner(value: object) -> str:
    text = normalize_text(value)
    if not text or text.lower() in {"--", "none", "nan", "null"}:
        return "Unassigned"
    return text


def unique_join(values: pd.Series, separator: str = "\n") -> str:
    cleaned_values = []
    seen = set()
    for value in values:
        text = normalize_text(value)
        if not text or text == "--":
            continue
        if text not in seen:
            seen.add(text)
            cleaned_values.append(text)
    return separator.join(cleaned_values)


def resolve_source_file() -> Path:
    if SOURCE_FILE.exists():
        return SOURCE_FILE

    candidates = [
        path
        for path in WORKSPACE_DIR.glob("rebootPendingList (3) 1 with owners*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("Could not find the pending reboot workbook with owners.")

    return max(candidates, key=lambda path: path.stat().st_mtime)


def main() -> None:
    source_file = resolve_source_file()
    dataframe = pd.read_excel(source_file)
    dataframe = normalize_columns(dataframe)

    required_columns = ["Owner", "Computer Name", "Owner Email", "IP Address"]
    missing_columns = [column for column in required_columns if column not in dataframe.columns]
    if missing_columns:
        raise KeyError(f"Missing required columns: {', '.join(missing_columns)}")

    dataframe = dataframe.loc[:, required_columns].copy()
    dataframe["Owner"] = dataframe["Owner"].map(normalize_owner)

    grouped = dataframe.groupby("Owner", sort=True, dropna=False)
    export_rows = []

    for owner_name, owner_frame in grouped:
        owner_name = normalize_owner(owner_name)
        export_rows.append(
            {
                "Owner": owner_name,
                "Owner Email": unique_join(owner_frame["Owner Email"]),
                "Computer Names": unique_join(owner_frame["Computer Name"], separator="\n"),
                "IP Addresses": unique_join(owner_frame["IP Address"], separator="\n"),
                "Server Count": int(len(owner_frame)),
            }
        )

    export_frame = pd.DataFrame(export_rows).sort_values(["Server Count", "Owner"], ascending=[False, True]).reset_index(drop=True)

    export_frame.to_csv(OUTPUT_FILE, index=False)
    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()