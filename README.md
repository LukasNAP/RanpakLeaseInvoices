# Ranpak lease invoices

Cleans Ranpak lease ("User fee billing") invoice Excel backups into the tab-delimited `.txt` that A+ loads from the F: drive (apedi-02), replacing the manual steps:

1. Remove the logo/header rows above the column headings.
2. Remove lines with a zero Period Amount.
3. Sort by Ship Date; machines shipped before 2022 use the Old Serial#.
4. Save as tab-delimited text.

Output for invoice 90210494 is byte-identical to the hand-made file.

## Usage

Drag one or more Excel backups onto `Clean Ranpak Invoices.bat`, or:

```
python ranpak_clean.py 90210494.xlsx [more.xlsx ...] [-o upload.txt]
```

Multiple invoices are combined into one file. Requires `openpyxl` (and `pypdf` for the PDF check).

## Checks

- **Stops, writes nothing:** layout changed, not "User fee billing", non-numeric amount, billed line without a ship date, or a total that doesn't match the invoice PDF ("Final amount") sitting in the same folder.
- **Writes, but warns:** duplicate serial # across invoices, pre-2022 machine with no Old Serial#, missing serial #, a field containing a line break or tab.

## Not yet

- Pre-check serial numbers against A+ before upload (waiting on which table/field the import matches on).
- Pull the Excel attachments straight from the invoice emails.
