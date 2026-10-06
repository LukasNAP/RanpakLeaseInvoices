"""Check serial numbers the same way the A+ import does, before the file is uploaded.

The ECS map "Ranpak_Invoice to Ranpak_Order" (per Trey) finds the customer for
each line by matching the serial to the Item Master's manufacturer item number
(ITMST.IMMFNO), then looks for:
  - Open Order Detail (ORDET) lines for that item -> customer / ship-to, or
  - Item Balance (ITBAL) with on-hand 1 and allocated 0 -> warehouse.
If neither finds anything, the import logs "No Customer Number found" and skips
the line. It takes the first row with no ORDER BY, so a serial that resolves to
more than one order (or, with no order, more than one warehouse) can land on an
unpredictable customer.

Reads DWStage on SQL03 (the A+ copy) with the user's Windows login.
"""
import pyodbc

SERVER = "SQL03.atlanticpkg.com"
DATABASE = "DWStage"

QUERY = """
SET NOCOUNT ON;
SELECT s.sn, RTRIM(i.IMITNO) AS item, 'order' AS src,
       RTRIM(d.ODCSNO) + '-' + RTRIM(d.ODSHP#) AS target
FROM #sn s
JOIN ITMST i ON RTRIM(i.IMMFNO) = s.sn
JOIN ORDET d ON d.ODITNO = i.IMITNO
UNION
SELECT s.sn, RTRIM(i.IMITNO), 'warehouse', 'warehouse ' + RTRIM(b.IBWHID)
FROM #sn s
JOIN ITMST i ON RTRIM(i.IMMFNO) = s.sn
JOIN ITBAL b ON b.IBITNO = i.IMITNO AND b.IBOHQ1 = 1 AND b.IBAQT1 = 0
UNION
SELECT s.sn, RTRIM(i.IMITNO), 'item only', NULL
FROM #sn s
JOIN ITMST i ON RTRIM(i.IMMFNO) = s.sn
"""


def _connect():
    # "SQL Server" ships with every Windows install; the newer driver is preferred when present.
    drivers = [d for d in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server")
               if d in pyodbc.drivers()]
    if not drivers:
        raise RuntimeError("no SQL Server ODBC driver installed")
    extra = "Encrypt=yes;TrustServerCertificate=yes;" if "ODBC Driver" in drivers[0] else ""
    return pyodbc.connect(
        f"DRIVER={{{drivers[0]}}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;{extra}",
        timeout=15,
    )


def check_serials(serials):
    """Return {serial: problem} for serials the A+ import would reject or could misroute."""
    serials = sorted({s for s in serials if s})
    with _connect() as conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE #sn (sn varchar(40) PRIMARY KEY)")
        cur.fast_executemany = True
        cur.executemany("INSERT INTO #sn VALUES (?)", [(s,) for s in serials])
        rows = cur.execute(QUERY).fetchall()

    items, targets = {}, {}
    for sn, item, src, target in rows:
        items.setdefault(sn, set()).add(item)
        if target:
            targets.setdefault(sn, set()).add(target)

    problems = {}
    for sn in serials:
        if sn not in items:
            problems[sn] = "not found in A+ Item Master (typo, or machine not set up yet)"
        elif sn not in targets:
            problems[sn] = (f"item {', '.join(sorted(items[sn]))} has no open order and isn't in a "
                            f"warehouse (machine may not be received into A+ yet)")
        else:
            # The order branch comes first in the import's UNION ALL, so an open order wins
            # over warehouse stock; only several candidates within the winning branch is ambiguous.
            orders = {t for t in targets[sn] if not t.startswith("warehouse")}
            candidates = orders or targets[sn]
            if len(candidates) > 1:
                problems[sn] = (f"matches more than one place in A+ ({'; '.join(sorted(candidates))}; "
                                f"items {', '.join(sorted(items[sn]))}) - the import may bill the wrong customer")
    return problems
