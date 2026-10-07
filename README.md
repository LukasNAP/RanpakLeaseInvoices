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

## Order file (new ZORHOF/ZORDOF map)

Alongside the upload `.txt`, the tool writes `ranpak_orders_<name>.txt`, the single order file agreed with Trey (2026-10-07) for a new Delta map into ZORHOF/ZORDOF/ZAOPRO. That map removes the manual offline order entry run. Pipe-delimited, CRLF, no header, one line per machine, grouped by order:

```
order_number|customer|ship_to|line_seq|item|description|serial|amount|invoice_number
RP001|213082|2|1|EQPRANUSER|STANDARD FILLPAK CONVERTER|11416700|20.00|90210494
```

- The customer lookup is the same as the current "Ranpak_Invoice to Ranpak_Order" map: an open order (`ORDET`) first, otherwise warehouse stock mapped to the branch's internal customer (e.g. 03 -> 33).
- One order per customer/ship-to, numbered `RP001`-`RP999` fresh each run, skipping any RP number still in live `ZORHOF`/`ZAOPRO` (a stuck order).
- `--override SERIAL=CUSTOMER-SHIPTO` settles a serial the lookup can't (e.g. one serial set up as two items).
- Checked against the 10/5 invoice: 210 of 211 lines match the orders A+ actually created; the other was the duplicate-item serial 10007655.

## Checks

- **Stops, writes nothing:** layout changed, not "User fee billing", non-numeric amount, billed line without a ship date, or a total that doesn't match the invoice PDF ("Final amount") sitting in the same folder.
- **Writes, but warns:** duplicate serial # across invoices, pre-2022 machine with no Old Serial#, missing serial #, a field containing a line break or tab.
- **A+ serial check (warns):** each serial is looked up the way the import does it (`aplus_check.py`): Item Master manufacturer item # (`ITMST.IMMFNO`), then Open Order Detail (`ORDET`) for the customer, or Item Balance (`ITBAL`) with on-hand 1 / allocated 0 for a warehouse. Flags serials the import would reject ("No Customer Number found") and serials that resolve to more than one place, since the import takes the first row with no ORDER BY.

## Not yet

- Pull the Excel attachments straight from the invoice emails.
