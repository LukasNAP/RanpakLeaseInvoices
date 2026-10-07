"""Build the single order file for the new ZORHOF/ZORDOF map (agreed with Trey 2026-10-07).

One file per upload, pipe-delimited, CRLF, no header, one line per machine,
grouped by order (one order per customer/ship-to, as the current maps do):

    order_number|customer|ship_to|line_seq|item|description|serial|amount|invoice_number

The tool does what the "Ranpak_Invoice to Ranpak_Order" map did (find the
customer for each serial) and numbers the orders RP001-RP999, skipping any RP
number still sitting in ZAOPRO or ZORHOF (a stuck order).
"""
from aplus_check import _connect

ITEM = "EQPRANUSER"

# Warehouse -> internal customer, copied from the CASE in the Ranpak_Invoice to
# Ranpak_Order map. Anything not listed passes through as the warehouse ID.
WAREHOUSE_CUSTOMER = {
    "01": "11", "03": "33", "05": "55", "06": "66", "07": "77", "08": "88",
    "09": "22", "11": "18", "12": "19", "13": "501113", "19": "190",
    "27": "210710", "28": "211744", "29": "29", "30": "212315", "33": "1333",
    "C7": "510116", "LV": "39",
}

LOOKUP = """
SET NOCOUNT ON;
SELECT s.sn, 1 AS pri, RTRIM(d.ODCSNO) AS cust, RTRIM(d.ODSHP#) AS shipto, NULL AS wh
FROM #sn s
JOIN ITMST i ON RTRIM(i.IMMFNO) = s.sn
JOIN ORDET d ON d.ODITNO = i.IMITNO
UNION
SELECT s.sn, 2, NULL, '1', RTRIM(b.IBWHID)
FROM #sn s
JOIN ITMST i ON RTRIM(i.IMMFNO) = s.sn
JOIN ITBAL b ON b.IBITNO = i.IMITNO AND b.IBOHQ1 = 1 AND b.IBAQT1 = 0
"""

# Live A+ staging tables, read through SQL03's APLUS linked server.
USED_RP = """
SELECT RTRIM(XHORNO) FROM OPENQUERY(APLUS, 'SELECT XHORNO FROM {lib}.ZORHOF WHERE XHORNO LIKE ''RP%''')
UNION
SELECT RTRIM(ZOORNO) FROM OPENQUERY(APLUS, 'SELECT ZOORNO FROM {lib}.ZAOPRO WHERE ZOORNO LIKE ''RP%''')
"""


def resolve_customers(serials):
    """Return {serial: (customer, ship_to)} using the same lookup as the ECS map.

    An open order wins over warehouse stock (the map's UNION ALL lists orders first).
    Serials with no match are left out; the caller reports them.
    """
    serials = sorted({s for s in serials if s})
    with _connect() as conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE #sn (sn varchar(40) PRIMARY KEY)")
        cur.fast_executemany = True
        cur.executemany("INSERT INTO #sn VALUES (?)", [(s,) for s in serials])
        rows = cur.execute(LOOKUP).fetchall()

    best = {}
    for sn, pri, cust, shipto, wh in sorted(rows, key=lambda r: (r[0], r[1], r[2] or "", r[3] or "", r[4] or "")):
        if sn in best:
            continue
        if pri == 2:
            cust = WAREHOUSE_CUSTOMER.get(wh, wh)
        best[sn] = (cust, shipto)
    return best


def used_rp_numbers(library="APLUSV8FAQ"):
    with _connect() as conn:
        return {r[0] for r in conn.cursor().execute(USED_RP.format(lib=library)).fetchall()}


def build_order_lines(lines, used, overrides=None):
    """lines: [{"invoice", "serial", "description", "amount"}] in upload order.

    overrides: {serial: (customer, ship_to)} for serials the lookup can't settle
    (e.g. one serial set up as two items). Returns (text_lines, orders, unmatched)
    where orders is [(rp, customer, ship_to, count, total)].
    """
    customers = resolve_customers(l["serial"] for l in lines)
    customers.update(overrides or {})
    groups, unmatched = {}, []
    for l in lines:
        target = customers.get(l["serial"])
        if target is None:
            unmatched.append(l)
            continue
        groups.setdefault(target, []).append(l)

    free = (f"RP{n:03d}" for n in range(1, 1000) if f"RP{n:03d}" not in used)
    out, orders = [], []
    for (cust, shipto), group in groups.items():
        try:
            rp = next(free)
        except StopIteration:
            raise RuntimeError("more than 999 orders in one upload - RP numbers exhausted")
        for seq, l in enumerate(group, 1):
            desc = l["description"].replace('"', "").replace("|", " ").strip()
            out.append("|".join([rp, cust, shipto, str(seq), ITEM, desc, l["serial"],
                                 f"{l['amount']:.2f}", l["invoice"]]))
        orders.append((rp, cust, shipto, len(group), sum(l["amount"] for l in group)))
    return out, orders, unmatched
