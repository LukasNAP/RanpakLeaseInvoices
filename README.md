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

Multiple invoices are combined into one file. Requires `openpyxl`, `pypdf` (PDF check) and `pyodbc` (A+ check, Windows login with read access to DWStage on SQL03). `--skip-aplus` turns the A+ check off.

## Checks

- **Stops, writes nothing:** layout changed, not "User fee billing", non-numeric amount, billed line without a ship date, or a total that doesn't match the invoice PDF ("Final amount") sitting in the same folder.
- **Writes, but warns:** duplicate serial # across invoices, pre-2022 machine with no Old Serial#, missing serial #, a field containing a line break or tab.
- **A+ serial check (warns):** each serial is looked up the way the import does it (`aplus_check.py`): Item Master manufacturer item # (`ITMST.IMMFNO`), then Open Order Detail (`ORDET`) for the customer, or Item Balance (`ITBAL`) with on-hand 1 / allocated 0 for a warehouse. Flags serials the import would reject ("No Customer Number found") and serials that resolve to more than one place, since the import takes the first row with no ORDER BY.

## Not yet

- Pull the Excel attachments straight from the invoice emails.
