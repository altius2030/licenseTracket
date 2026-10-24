# -*- coding: utf-8 -*-
"""Smoke test: toolbar layout, renewal logic."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import datetime as dt

import main as m
from database import Database

TEST_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_renew.db")
if os.path.exists(TEST_DB):
    os.remove(TEST_DB)

db = Database(TEST_DB)
app = m.LicenseTrackerApp()
app.withdraw()
app.db = db  # use the test database instead of the production one

today = dt.date.today()

# --- add_months ---
assert m.add_months(dt.date(2026, 1, 31), 1) == dt.date(2026, 2, 28), "clamp Feb"
assert m.add_months(dt.date(2026, 11, 20), 3) == dt.date(2027, 2, 20), "cross year"
assert m.add_months(dt.date(2026, 1, 15), 12) == dt.date(2027, 1, 15), "year"
print("ADD MONTHS OK")

# --- parse_date_any ---
assert m.parse_date_any("2026-05-01") == dt.datetime(2026, 5, 1)
assert m.parse_date_any("01.05.2026") == dt.datetime(2026, 5, 1)
assert m.parse_date_any("2026-05-01T10:30:00") == dt.datetime(2026, 5, 1, 10, 30)
assert m.parse_date_any("") is None
assert m.parse_date_any(None) is None
assert m.parse_date_any("мусор") is None
print("PARSE DATE ANY OK")

# --- Renewal: expired record renews from today ---
lid1 = db.create_license({'sheet_name': 'GBS', 'category': 'GBS-market', 'full_name': 'Просрочка',
                          'id_expiry': '2026-01-01'})
lic = db.get_license(lid1)
app._apply_renewal([lic], 1, 0, dlg=None)
new_val = db.get_license(lid1)['id_expiry']
expected = m.add_months(today, 1)
assert new_val == expected.isoformat(), f"expired -> from today: {new_val} != {expected}"
print("RENEW EXPIRED FROM TODAY OK:", new_val)

# --- Renewal: valid record extends from current expiry ---
future = today + dt.timedelta(days=100)
lid2 = db.create_license({'sheet_name': 'GBS', 'category': 'GBS-market', 'full_name': 'Активная',
                          'id_expiry': future.isoformat()})
lic = db.get_license(lid2)
app._apply_renewal([lic], 3, 0, dlg=None)
new_val2 = db.get_license(lid2)['id_expiry']
expected2 = m.add_months(future, 3)
assert new_val2 == expected2.isoformat(), f"valid -> extend: {new_val2} != {expected2}"
print("RENEW VALID EXTENDS OK:", new_val2)

# --- Renewal: missing expiry -> from today ---
lid3 = db.create_license({'sheet_name': 'GBS', 'category': 'GBS-market', 'full_name': 'Без срока'})
lic = db.get_license(lid3)
app._apply_renewal([lic], 0, 45, dlg=None)  # custom: 45 days
new_val3 = db.get_license(lid3)['id_expiry']
expected3 = today + dt.timedelta(days=45)
assert new_val3 == expected3.isoformat(), f"missing -> today+45: {new_val3} != {expected3}"
print("RENEW MISSING WITH CUSTOM DAYS OK:", new_val3)

# --- Renewal: multiple at once ---
app._apply_renewal([db.get_license(i) for i in (lid1, lid2, lid3)], 6, 0, dlg=None)
vals = [db.get_license(i)['id_expiry'] for i in (lid1, lid2, lid3)]
print("BULK RENEW:", vals)
assert vals[0] == m.add_months(expected, 6).isoformat()
assert vals[1] == m.add_months(expected2, 6).isoformat()
assert vals[2] == m.add_months(expected3, 6).isoformat()
print("BULK RENEW OK")

# --- Toolbar widgets present ---
assert hasattr(app, 'search_entry') and hasattr(app, 'theme_btn')
app.destroy()
os.remove(TEST_DB)
print("ALL RENEW TESTS PASSED")
