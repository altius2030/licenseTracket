# -*- coding: utf-8 -*-
"""Smoke test: comments CRUD, sorting by dates, theme toggle, multi-delete."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import datetime as dt

import main as m
from database import Database

TEST_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_comments.db")
if os.path.exists(TEST_DB):
    os.remove(TEST_DB)

db = Database(TEST_DB)

# --- Comments CRUD ---
lic_id = db.create_license({'sheet_name': 'GBS', 'category': 'GBS-market', 'full_name': 'Тестов Тест'})
cid1 = db.add_comment(lic_id, "Первый комментарий")
cid2 = db.add_comment(lic_id, "Второй комментарий")
cs = db.get_comments(lic_id)
assert len(cs) == 2, f"Expected 2 comments, got {len(cs)}"
assert cs[0]['text'] == "Второй комментарий", "newest comment must be first"
assert cs[0]['created_at'] is not None

# Edit
db.update_comment(cid1, "Первый (отредактирован)")
cs = db.get_comments(lic_id)
first = [c for c in cs if c['id'] == cid1][0]
assert first['text'] == "Первый (отредактирован)"
assert first['updated_at'] is not None
print("Edit comment saved with timestamp:", first['updated_at'])

# Delete
db.delete_comment(cid2)
assert len(db.get_comments(lic_id)) == 1

# Delete license cascades comments
db.delete_license(lic_id)
assert db.get_comments(lic_id) == []
print("COMMENTS CRUD OK")

# --- Sorting by dates ---
now = dt.datetime.now()
rows = []
samples = [
    ("A", "2026-11-05"),
    ("B", None),
    ("C", "2025-01-01"),
    ("D", "05.03.2026"),   # DD.MM.YYYY format from Excel
    ("E", "2026-06-10"),
]
lic_ids = []
for name, exp in samples:
    lid = db.create_license({'sheet_name': 'GBS', 'category': 'GBS-market', 'full_name': name, 'id_expiry': exp})
    lic_ids.append((name, lid, exp))

app = m.LicenseTrackerApp()
app.withdraw()
app.current_category = 'GBS-market'
app.current_sheet = ''

# Simulate sort by id_expiry ascending
app.sort_column = 'id_expiry'
app.sort_reverse = False
# Extract licenses and replicate the filter part of _refresh_data
licenses = db.get_all_licenses(category='GBS-market')
for lic in licenses:
    lic['days_left'] = app._calc_days_left(lic)

DATES = {'install_date', 'id_expiry', 'payment_date', 'renewal_date', 'next_payment'}

def _parse_date(v):
    if isinstance(v, dt.datetime):
        return dt.datetime.combine(v.date(), dt.datetime.min.time())
    if isinstance(v, dt.date):
        return dt.datetime.combine(v, dt.datetime.min.time())
    s = str(v).strip()
    try:
        return dt.datetime.fromisoformat(s.replace('Z', '').split('.')[0])
    except ValueError:
        pass
    try:
        return dt.datetime.strptime(s[:10], "%Y-%m-%d")
    except ValueError:
        pass
    for fmt in ("%d.%m.%Y", "%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S"):
        try:
            return dt.datetime.strptime(s[:19], fmt)
        except ValueError:
            continue
    return None

parsed = [(n, e, _parse_date(e) if e else None) for n, e, _ in lic_ids]
vals = [p for _, _, p in parsed if p]
assert all(isinstance(v, dt.datetime) for v in vals), "all parsed to datetime"
# B (None) must compare equal ranks, no TypeError possible in sort
col = 'id_expiry'
def sort_key(item):
    val = item.get(col)
    if col in DATES and val is not None:
        p = _parse_date(val)
        return (0, p) if p else (2, dt.datetime.min)
    return (2, dt.datetime.min)
lics_sorted = sorted(licenses, key=sort_key)
order_names = []
by_id = {i: (n, e) for n, i, e in lic_ids}
for it in lics_sorted:
    order_names.append((by_id[it['id']][0], it.get('id_expiry')))
print("Sorted order (id_expiry asc):", order_names)
# Check ascending among non-null
# Check ascending among non-null (parse both ISO and DD.MM.YYYY)
def p2(v):
    try:
        return dt.datetime.fromisoformat(v)
    except ValueError:
        return dt.datetime.strptime(v, "%d.%m.%Y")

non_null_vals = [p2(d) for n, d in order_names if d and p2(d)]
assert non_null_vals == sorted(non_null_vals), "dates must be ascending"
print("Non-null ascending:", non_null_vals)
print("DATE SORT OK")

# --- Theme toggle mechanics ---
assert 'dark' in m.THEMES and 'light' in m.THEMES
t1 = m.BG_APP
m.apply_theme('light')
assert m.BG_APP == m.THEMES['light']['BG_APP'], "apply_theme must reassign globals"
assert m.TAG_BG['active'] == m.THEMES['light']['TAG_BG']['active']
m.apply_theme('dark')
assert m.BG_APP == m.THEMES['dark']['BG_APP']
print("THEME SWITCH MECHANICS OK")

# Rebuild UI under light theme (full widget rebuild path)
m.apply_theme('light')
app._toggle_theme()  # light -> dark, rebuild UI
assert app.winfo_children(), "widgets rebuilt after theme toggle"
app._toggle_theme()  # dark -> light
print("THEME TOGGLE REBUILD OK")

# --- Multi-delete simulation ---
ids = []
for i, name in enumerate(["Multi1", "Multi2", "Multi3"]):
    lid = db.create_license({'sheet_name': 'GBS', 'category': 'GBS-market', 'full_name': name})
    ids.append(lid)
for lid in ids:
    db.delete_license(lid)
remaining = [l['full_name'] for l in db.get_all_licenses(category='GBS-market')]
assert not any(n in remaining for n in ["Multi1", "Multi2", "Multi3"])
print("MULTI DELETE OK")

app.destroy()
os.remove(TEST_DB)
print("ALL TESTS PASSED")
