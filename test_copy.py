# -*- coding: utf-8 -*-
"""Copy-from-card feature tests."""
import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import main as m

# 1) Pure formatting (the user's example)
sample = {
    'full_name': 'Джаикбаева Камила',
    'email': 'Dzhaikbaeva@bk.ru',
    'store_name': 'Магазин одежды',
    'phone': '8747 346 54 36',
    'gbs_id': '6:E61991B5D4BB4173A3301A899A31EF77',
}
expected = ("Джаикбаева Камила\n"
            "Dzhaikbaeva@bk.ru\n"
            "Магазин одежды\n"
            "8747 346 54 36\n"
            "6:E61991B5D4BB4173A3301A899A31EF77")
result = m.format_card_copy_text(sample)
assert result == expected, f"format mismatch:\n{result!r}\n!=\n{expected!r}"
print("FORMAT OK")

# Empty / missing values are skipped, empty fields don't leave blank lines
partial = {'full_name': 'Иванов Олег', 'email': '', 'store_name': 'Автозапчасти',
           'phone': None, 'license_key': 'ABC-123'}
assert m.format_card_copy_text(partial) == "Иванов Олег\nАвтозапчасти\nABC-123"
assert m.format_card_copy_text({}) == ""
assert m.format_card_copy_text(None) == ""
print("PARTIAL/EMPTY OK")

app = m.LicenseTrackerApp()
app.geometry("1200x760+80+40")
app.deiconify()
app.update()
time.sleep(0.3)
app.update()

# 2) Card dialog: button copies current entry values (including unsaved edits)
lic = app.db.get_all_licenses(category=app.current_category)[0]
dlg = m.LicenseCardDialog(app, app.db, license_data=lic,
                          sheet_name=lic.get('sheet_name'),
                          category=lic.get('category'))
dlg.update()
assert dlg.copy_btn.cget("text").startswith("📋")
dlg.copy_btn.invoke()
app.update()
copied = app.clipboard_get()
expected_dialog = m.format_card_copy_text(dlg._collect_data())
assert copied == expected_dialog, f"clipboard:\n{copied!r}\nexpected:\n{expected_dialog!r}"
assert dlg.copy_btn.cget("text").startswith("✓")
print("DIALOG COPY OK ->", copied.replace("\n", " | "))

# unsaved edit is included
dlg.entries['full_name'].delete(0, 'end')
dlg.entries['full_name'].insert(0, 'Тестовое Имя')
dlg.copy_btn.invoke()
app.update()
assert app.clipboard_get().splitlines()[0] == 'Тестовое Имя'
print("UNSAVED EDITS COPIED OK")

# Ctrl+Shift+C shortcut is registered (key events can't be synthesized
# reliably without an OS-focused window, so verify the binding itself)
assert dlg.bind("<Control-Shift-C>"), "Ctrl+Shift+C must be bound in the card"
print("SHORTCUT BINDING OK")

dlg.destroy()
app.update()

# 3) Context menu action copies selected row(s)
app.tree.selection_set(app.tree.get_children()[0])
app.update()
app._on_copy_card_data()
app.update()
first_id = int(app.tree.selection()[0])
row = app.db.get_license(first_id)
copied_row = app.clipboard_get()
expected_row = m.format_card_copy_text(row)
assert copied_row == expected_row, f"row clipboard mismatch:\n{copied_row!r}\n{expected_row!r}"
print("CONTEXT MENU COPY OK ->", copied_row.replace("\n", " | "))

# multi-selection -> blocks separated by a blank line
children = app.tree.get_children()[:2]
app.tree.selection_set(children)
app.update()
app._on_copy_card_data()
app.update()
multi = app.clipboard_get()
assert "\n\n" in multi, "two selected rows must be separated by a blank line"
print("MULTI-SELECT COPY OK")

app.destroy()
print("ALL COPY TESTS PASSED")
