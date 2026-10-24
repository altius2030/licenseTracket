"""Smoke test for the license card dialog UI."""
import sys
sys.path.insert(0, r'C:\Users\user\Desktop\Лицензии\LicenseTracker')
import main as m
from database import Database

app = m.LicenseTrackerApp()
app.update_idletasks()

# Open edit dialog for first record
lic = app.db.get_all_licenses(category='GBS-market')[0]
dlg = m.LicenseCardDialog(app, app.db, lic, category='GBS-market')
app.update_idletasks()

# Check all fields loaded
loaded = 0
for fname, entry in dlg.entries.items():
    val = getattr(entry, 'get', None)()
    if isinstance(entry, __import__('customtkinter').CTkTextbox):
        val = entry.get("1.0", "end-1c")
    if val:
        loaded += 1
print("Fields loaded:", loaded)

# Check date formatting in the entry
exp = dlg.entries['next_payment'].get()
print("next_payment display:", repr(exp))
assert "/" not in exp, "date should be DD.MM.YYYY"

# Check status section (automatic status only, no manual override)
print("auto status:", dlg.auto_status_label.cget("text"))
assert not hasattr(dlg, "manual_status_seg"), "manual override must be removed from the card"

# Test collect+save roundtrip
data = dlg._collect_data()
assert 'manual_status' not in data, "card must not send manual status anymore"

dlg.destroy()
app.destroy()
print("DIALOG TEST PASSED")
