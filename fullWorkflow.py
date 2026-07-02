from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd


# --- Utility helpers -----------------------------------------------------

TAG_PATTERN = re.compile(r"<([^<>]+)>")


def choose_file_dialog(prompt: str, filetypes=None) -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        selected = filedialog.askopenfilename(title=prompt, filetypes=filetypes)
        return Path(selected) if selected else None
    except Exception:
        # Fallback to console input
        print(f"{prompt} - please paste the full path (or leave empty to cancel):")
        line = input().strip()
        return Path(line) if line else None


def normalize_header(name: str) -> str:
    return str(name).replace("\ufeff", "").strip().strip('"')


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\ufeff", "").strip().strip("'\"")


def normalize_owner(value: object) -> str:
    txt = normalize_text(value)
    if not txt or txt.lower() in {"--", "none", "nan", "null"}:
        return "Unassigned"
    return txt


def normalize_key(s: str) -> str:
    # normalize names for fuzzy matches: lowercase, remove punctuation/whitespace
    return re.sub(r"[^0-9a-z]", "", str(s).lower())


def normalize_identifier(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_key(s))


def split_ip_addresses(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []

    parts = [part.strip() for part in re.split(r"[;,]\s*", text) if part.strip()]
    return parts or [text]


# --- vSphere cleaning (case-insensitive tag merge) ---------------------


def parse_tags_value(raw_value: object) -> dict[str, str]:
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
        key_norm = key.strip().lower()
        value = value.strip()
        if key_norm:
            tags[key_norm] = value
    return tags


def row_has_valid_tags(raw_value: object) -> bool:
    text = normalize_text(raw_value)
    if not text or text.lower() == "none":
        return False

    tags = parse_tags_value(text)
    if not tags:
        return False

    # Drop the row if any tag entry is effectively a none placeholder.
    for key, value in tags.items():
        if not key.strip() or not value.strip() or key.strip().lower() == "none" or value.strip().lower() == "none":
            return False

    return True


PREFERRED_COLUMNS = [
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


def clean_vsphere_inventory(source: Path, output_dir: Path) -> Path:
    # read CSV or Excel
    if source.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(source, dtype=str)
    else:
        df = pd.read_csv(source, dtype=str)

    # normalize source headers so quoted names like "Name" become Name
    df = df.rename(columns={c: normalize_header(c) for c in df.columns})

    # ensure 'vSphere Tag' exists
    if "vSphere Tag" not in df.columns:
        raise KeyError("Source inventory does not contain a 'vSphere Tag' column.")

    # parse tags
    non_empty = df["vSphere Tag"].apply(row_has_valid_tags)
    tag_rows = df.loc[non_empty, "vSphere Tag"].apply(parse_tags_value)
    df = df.loc[non_empty].copy().reset_index(drop=True)
    tag_rows = tag_rows.loc[non_empty].reset_index(drop=True)

    tag_columns = pd.DataFrame(list(tag_rows), index=df.index)

    # Map lowercased tag keys to canonical column names (case-insensitive)
    mapping: dict[str, str] = {}
    for col in tag_columns.columns:
        match = next((pc for pc in PREFERRED_COLUMNS if pc.lower() == col.lower()), None)
        if match:
            mapping[col] = match
        else:
            mapping[col] = col.strip().title()
    if mapping:
        tag_columns = tag_columns.rename(columns=mapping)

    if tag_columns.columns.duplicated().any():
        tag_columns = tag_columns.groupby(tag_columns.columns, axis=1).first()

    combined = pd.concat([df.drop(columns=["vSphere Tag"]), tag_columns], axis=1)

    # reorder
    ordered = [c for c in PREFERRED_COLUMNS if c in combined.columns]
    ordered.extend(c for c in combined.columns if c not in ordered)
    combined = combined.loc[:, ordered]

    out_name = f"{source.stem} cleaned.xlsx"
    out_path = output_dir / out_name
    combined.to_excel(out_path, index=False)
    return out_path


# --- Build inventory lookup for matching --------------------------------


def build_inventory_lookup(cleaned_df: pd.DataFrame) -> dict:
    # We'll build four maps: exact (name, ip) -> custodian, name -> custodian,
    # normalized name -> custodian, ip -> custodian
    exact_lookup: dict[tuple[str, str], str] = {}
    name_map: dict[str, str] = {}
    ip_map: dict[str, pd.Series] = {}
    normalized_map: dict[str, str] = {}

    # determine which columns to use
    name_col = next((c for c in ("Name", "Computer Name", "Computer" ) if c in cleaned_df.columns), None)
    ip_col = next((c for c in ("IP Address", "IP", "IP Address " ) if c in cleaned_df.columns), None)
    custodian_col = next((c for c in ("Asset Custodian", "Custodian") if c in cleaned_df.columns), None)

    if name_col is None or ip_col is None or custodian_col is None:
        raise KeyError("Cleaned inventory must contain Name, IP Address, and Asset Custodian columns.")

    for _, row in cleaned_df.iterrows():
        custodian = normalize_text(row.get(custodian_col, ""))
        if not custodian:
            continue

        name_val = normalize_text(row.get(name_col, ""))
        if not name_val:
            continue

        name_key = normalize_key(name_val)
        normalized_name_key = normalize_key(name_val)

        if name_key and name_key not in name_map:
            name_map[name_key] = custodian

        if normalized_name_key and normalized_name_key not in normalized_map:
            normalized_map[normalized_name_key] = custodian

        for ip_address in split_ip_addresses(row.get(ip_col, "")):
            normalized_ip = normalize_key(ip_address)
            if normalized_ip and normalized_ip not in ip_map:
                ip_map[normalized_ip] = custodian

            key = (name_key, normalized_ip)
            if key not in exact_lookup:
                exact_lookup[key] = custodian

    return {"exact": exact_lookup, "name": name_map, "normalized_name": normalized_map, "ip": ip_map}


# --- Fill pending reboot owners -----------------------------------------


def fill_pending_owners(pending_df: pd.DataFrame, lookup: dict) -> pd.DataFrame:
    df = pending_df.copy()
    # normalize column names
    rename = {c: normalize_header(c) for c in df.columns}
    df = df.rename(columns=rename)

    # find candidate columns
    comp_col = next((c for c in ("Computer Name", "Name", "Computer") if c in df.columns), None)
    ip_col = next((c for c in ("IP Address", "IP") if c in df.columns), None)
    owner_col = next((c for c in ("Owner", "owner") if c in df.columns), "Owner")

    if comp_col is None:
        raise KeyError("Pending reboot file must contain a 'Computer Name' or 'Name' column")

    filled = []
    for _, row in df.iterrows():
        owner_val = normalize_owner(row.get(owner_col, ""))
        if owner_val != "Unassigned":
            filled.append(owner_val)
            continue

        # attempt exact computer name match
        comp_val = normalize_text(row.get(comp_col, ""))
        ip_val = normalize_text(row.get(ip_col, "")) if ip_col else ""
        candidate = None
        name_key = normalize_key(comp_val)
        if name_key:
            for ip_address in split_ip_addresses(ip_val):
                normalized_ip = normalize_key(ip_address)
                candidate = lookup["exact"].get((name_key, normalized_ip))
                if candidate:
                    break

        if not candidate:
            for ip_address in split_ip_addresses(ip_val):
                candidate = lookup["ip"].get(normalize_key(ip_address))
                if candidate:
                    break

        if not candidate and name_key:
            candidate = lookup["name"].get(name_key)

        if not candidate and name_key:
            candidate = lookup["normalized_name"].get(normalize_identifier(comp_val))

        if candidate:
            filled.append(candidate)
        else:
            filled.append("Unassigned")

    df[owner_col] = filled
    # ensure Owner Email column exists
    if "Owner Email" not in df.columns:
        df["Owner Email"] = ""
    return df


# --- Outlook email resolution (best-effort) -----------------------------


def resolve_email_outlook(name: str, cache: dict) -> Optional[str]:
    if not name or name.lower() == "unassigned":
        return None
    key = name.strip().lower()
    if key in cache:
        return cache[key]
    try:
        import win32com.client

        session = win32com.client.Dispatch("MAPI")
        recip = session.CreateRecipient(name)
        recip.Resolve()
        if recip.Resolved:
            exch = recip.GetExchangeUser()
            if exch is not None:
                address = getattr(exch, "PrimarySmtpAddress", None)
                if address:
                    cache[key] = address
                    return address
        # fallback: try AddressEntry
        cache[key] = None
        return None
    except Exception:
        cache[key] = None
        return None


# --- Export per-owner CSV ------------------------------------------------


def unique_join(values: pd.Series, separator: str = "\n") -> str:
    cleaned = []
    seen = set()
    for v in values:
        t = normalize_text(v)
        if not t or t == "--":
            continue
        if t not in seen:
            seen.add(t)
            cleaned.append(t)
    return separator.join(cleaned)


def normalize_email_value(value: object) -> str:
    text = normalize_text(value)
    return text if text else "not found"


def export_per_owner_csv(pending_df: pd.DataFrame, output_csv: Path) -> Path:
    # normalize headers
    df = pending_df.copy()
    df = df.rename(columns={c: normalize_header(c) for c in df.columns})

    required = ["Owner", "Computer Name", "Owner Email", "IP Address"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Missing required columns for export: {', '.join(missing)}")

    df = df.loc[:, required].copy()
    df["Owner"] = df["Owner"].map(normalize_owner)

    rows = []
    for owner, group in df.groupby("Owner", sort=True, dropna=False):
        owner_name = normalize_owner(owner)
        rows.append(
            {
                "Owner": owner_name,
                "Owner Email": unique_join(group["Owner Email"]) or "not found",
                "Computer Names": unique_join(group["Computer Name"], separator="\n"),
                "IP Addresses": unique_join(group["IP Address"], separator="\n"),
                "Server Count": int(len(group)),
            }
        )
    out = pd.DataFrame(rows).sort_values(["Server Count", "Owner"], ascending=[False, True]).reset_index(drop=True)
    out.to_csv(output_csv, index=False)
    return output_csv


# --- Populate owner template (basic mapping) ----------------------------


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


def build_template_from_inventory(cleaned_df: pd.DataFrame, output_path: Path, email_cache: dict) -> Path:
    rows = []
    # find candidate columns
    name_col = next((c for c in ("Name", '"Name"', "Computer Name") if c in cleaned_df.columns), None)
    owner_col = next((c for c in ("Asset Custodian", "Asset custodian", "Asset Owner", "Asset owner", "Owner") if c in cleaned_df.columns), None)
    manager_col = next((c for c in ("Asset Owner", "Asset owner", "Owner") if c in cleaned_df.columns), None)
    criticality_col = next((c for c in ("Criticality of the VM", "Criticality", "SEARCH_TAG") if c in cleaned_df.columns), None)
    email_col = next((c for c in ("Asset Owner Email", "Owner Email") if c in cleaned_df.columns), None)

    dep_col = next((c for c in ("Department / Sections", "Department", "Department/Sections") if c in cleaned_df.columns), None)

    for _, r in cleaned_df.iterrows():
        comp = normalize_text(r.get(name_col, "")) if name_col else ""
        owner = normalize_text(r.get(owner_col, "")) if owner_col else ""
        # try email from row first
        owner_email = normalize_text(r.get(email_col, "")) if email_col else ""
        if not owner_email and owner:
            owner_email = resolve_email_outlook(owner, email_cache) or ""
        notes_parts = []
        if dep_col:
            depv = normalize_text(r.get(dep_col, ""))
            if depv:
                notes_parts.append(f"Department / Sections: {depv}")
        manager_value = normalize_text(r.get(manager_col, "")) if manager_col else ""
        if manager_value:
            notes_parts.append(f"Manager: {manager_value}")
        if len(notes_parts) == 2:
            notes = ". ".join(notes_parts)
        elif notes_parts:
            notes = notes_parts[0]
        else:
            notes = ""
        rows.append(
            {
                "COMPUTER_NAME": comp,
                "DOMAIN_NAME": "",
                "CUSTOMER_NAME": "",
                "OWNER": owner,
                "LOCATION": normalize_text(r.get("Location", "")) if "Location" in cleaned_df.columns else "",
                "SEARCH_TAG": normalize_text(r.get(criticality_col, "")) if criticality_col else "",
                "NOTES": notes,
                "PRODUCT_NUMBER": "",
                "SHIPPING_DATE": "",
                "WARRANTY_EXPIRY_DATE": "",
                "OWNER_EMAIL_ID": normalize_email_value(owner_email),
            }
        )

    out = pd.DataFrame(rows, columns=TEMPLATE_COLUMNS)
    out.to_excel(output_path, index=False)
    return output_path


# --- Main workflow ------------------------------------------------------


def main() -> None:
    cwd = Path.cwd()
    print("Select the vSphere inventory file (CSV or Excel)")
    inventory_path = choose_file_dialog("Input inventory file", filetypes=[("CSV and Excel", "*.csv *.xlsx *.xls")])
    if not inventory_path:
        print("Inventory selection cancelled. Exiting.")
        return

    print("Select the pending reboot file (CSV or Excel)")
    pending_path = choose_file_dialog("Input pending reboot file", filetypes=[("CSV and Excel", "*.csv *.xlsx *.xls")])
    if not pending_path:
        print("Pending reboot selection cancelled. Exiting.")
        return

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = cwd / f"full_workflow_output_{timestamp}"
    out_dir.mkdir(exist_ok=True)

    # Step 1: clean inventory
    print("Cleaning vSphere inventory...")
    cleaned_inv_path = clean_vsphere_inventory(inventory_path, out_dir)
    cleaned_df = pd.read_excel(cleaned_inv_path, dtype=str).fillna("")

    # Step 2: build lookup and fill pending owners
    print("Building inventory lookup...")
    lookup = build_inventory_lookup(cleaned_df)

    print("Reading pending reboot file and filling missing owners...")
    if pending_path.suffix.lower() in {".xlsx", ".xls"}:
        pending_df = pd.read_excel(pending_path, dtype=str)
    else:
        pending_df = pd.read_csv(pending_path, dtype=str)
    pending_df = pending_df.fillna("")

    filled_df = fill_pending_owners(pending_df, lookup)

    filled_out_path = out_dir / f"{pending_path.stem} with owners.xlsx"
    filled_df.to_excel(filled_out_path, index=False)

    # Step 3: resolve owner emails
    print("Resolving owner emails via Outlook (best-effort)...")
    email_cache: dict = {}
    owner_col = "Owner"
    if owner_col not in filled_df.columns:
        # attempt to find correct column
        owner_col = next((c for c in filled_df.columns if c.lower() == "owner"), "Owner")
    if "Owner Email" not in filled_df.columns:
        filled_df["Owner Email"] = ""

    for i, row in filled_df.iterrows():
        owner = normalize_owner(row.get(owner_col, ""))
        if owner == "Unassigned":
            continue
        email = normalize_text(row.get("Owner Email", ""))
        if email and email != "--":
            continue
        resolved = resolve_email_outlook(owner, email_cache)
        filled_df.at[i, "Owner Email"] = resolved if resolved else "not found"

    filled_df["Owner Email"] = filled_df["Owner Email"].map(normalize_email_value)

    # overwrite the filled output file with emails added
    filled_df.to_excel(filled_out_path, index=False)

    # Step 4: export per-owner CSV
    print("Exporting per-owner CSV...")
    csv_out = out_dir / f"{pending_path.stem} owners listed.csv"
    try:
        export_per_owner_csv(filled_df, csv_out)
    except Exception as exc:
        print(f"Failed to export per-owner CSV: {exc}")

    # Step 5: build template workbook from cleaned inventory
    print("Building owner template workbook...")
    template_out = out_dir / f"{inventory_path.stem} owner_template.xlsx"
    try:
        build_template_from_inventory(cleaned_df, template_out, email_cache)
    except Exception as exc:
        print(f"Failed to build template workbook: {exc}")

    # Step 6: save cleaned inventory (already saved) and summary
    print("Workflow complete. Outputs written to:")
    for p in sorted(out_dir.iterdir()):
        print(f" - {p.name}")

    print("Summary:")
    print(f" Cleaned inventory: {cleaned_inv_path.name}")
    print(f" Filled pending reboot: {filled_out_path.name}")
    print(f" Per-owner CSV: {csv_out.name}")
    print(f" Owner template: {template_out.name}")


if __name__ == "__main__":
    main()
