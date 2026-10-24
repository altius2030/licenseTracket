"""Quick UI smoke test: create the app, load data, simulate filters, destroy."""
import sys
sys.path.insert(0, r'C:\Users\user\Desktop\Лицензии\LicenseTracker')
import main as m
import customtkinter as ctk

app = m.LicenseTrackerApp()

# Simulate category switching
for cat in m.CATEGORIES:
    app._on_category_change(cat)
    app.update_idletasks()
    rows = len(app.tree.get_children())
    print(f"Category '{cat}': {rows} rows, chips total={app.stat_chips['total'].cget('text')}")

# Back to first
app._on_category_change("GBS-market")
app.update_idletasks()

# Simulate days filter
for f in ["0-3 дня", "5 дней", "Неделя (7 дней)", "10 дней", "Просроченные", "Все"]:
    app.days_filter_var.set(f)
    app._on_filter_change()
    app.update_idletasks()
    rows = len(app.tree.get_children())
    print(f"Filter '{f}': {rows} rows")

# Simulate status filter
for f in ["Активна", "Истекла", "Ручная", "Все"]:
    app.status_filter_var.set(f)
    app._on_filter_change()
    app.update_idletasks()
    print(f"Status '{f}': {len(app.tree.get_children())} rows")

# Reset filters
app._reset_filters()
app.update_idletasks()
print(f"Reset: {len(app.tree.get_children())} rows")

# Sheet switching
app._load_sheets()
app.update_idletasks()
print("Sheet seg values:", app.sheet_seg.cget("values"))
app._on_sheet_change("Все")

app.update_idletasks()
app.destroy()
print("UI SMOKE TEST PASSED")
