import csv

import openpyxl
import win32com.client

# Open Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Load Excel
wb = openpyxl.load_workbook("names.xlsx")
ws = wb.active

rows = []

# Add Email column if needed
if ws["B1"].value != "Email":
    ws["B1"] = "Email"

for row in range(2, ws.max_row + 1):
    name = ws[f"A{row}"].value

    if not name:
        continue

    recipient = namespace.CreateRecipient(name)
    recipient.Resolve()

    email = ""

    if recipient.Resolved:
        try:
            exch_user = recipient.AddressEntry.GetExchangeUser()
            if exch_user:
                email = exch_user.PrimarySmtpAddress
            else:
                email = recipient.Address
        except Exception:
            email = recipient.Address
    else:
        email = "Not Found"

    ws[f"B{row}"] = email
    rows.append([name, email])
    print(f"{name} -> {email}")

with open("names_with_emails.csv", "w", newline="", encoding="utf-8-sig") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(["Name", "Email"])
    writer.writerows(rows)

print("Done!")