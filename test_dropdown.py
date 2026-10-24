# -*- coding: utf-8 -*-
"""Dropdowns: real left-click path (CTkButton._on_release), open above,
selection, click-outside close, scrollable lists, theme/category/sheet toggles."""
import sys, os, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
import main as m

errors = []


def hook(exc, val, tb):
    import traceback
    errors.append("".join(traceback.format_exception(exc, val, tb)))
    traceback.print_exception(exc, val, tb)


app = m.LicenseTrackerApp()
app.report_callback_exception = hook
app.geometry("1200x760+80+40")
app.deiconify()
app.update()
time.sleep(0.2)
app.update()


def real_click(widget):
    """Exactly what a physical left click does inside CTkButton."""
    widget._mouse_inside = True
    widget._on_release()


def open_check(btn, label):
    real_click(btn)
    app.update()
    time.sleep(0.25)
    app.update()
    popup = btn._popup
    assert popup is not None and popup.winfo_exists(), f"{label}: list did not open on click"
    assert popup.winfo_viewable() == 1, f"{label}: list not visible"
    btn_top = btn.winfo_rooty()
    btn_bottom = btn_top + btn.winfo_height()
    pop_top = popup.winfo_rooty()
    pop_bottom = pop_top + popup.winfo_height()
    if pop_bottom <= btn_top + 6:
        print(f"{label}: OK above (popup {pop_top}..{pop_bottom}, button {btn_top})")
    else:
        assert pop_top >= btn_bottom - 6, (
            f"{label}: popup must be above or below the button "
            f"(popup {pop_top}..{pop_bottom}, button {btn_top}..{btn_bottom})")
        print(f"{label}: OK below (no room above)")
    return popup


# Regression guard: the click must open the list, not run the filter callback
btn = app.days_filter_menu
assert btn._command == btn._open_menu, \
    f"left click runs {btn._command}, must run the dropdown opener"
print("click command wiring: OK")

open_check(btn, "days filter (7 items)")
btn._choose("10 дней")
app.update()
assert app.days_filter_var.get() == "10 дней" and btn._popup is None
print("days choice: OK")

# toggle: click the button again closes the list
real_click(btn); app.update(); time.sleep(0.2); app.update()
assert btn._popup is not None
real_click(btn); app.update()
assert btn._popup is None, "second click on button must close the list"
print("toggle: OK")

# click outside closes
real_click(btn); app.update(); time.sleep(0.2); app.update()
assert btn._popup is not None
app.tree.event_generate("<Button-1>", x=5, y=5)
app.update(); time.sleep(0.15); app.update()
assert btn._popup is None, "click outside must close the list"
print("click-outside close: OK")

btn._choose("Все")
app.update()

open_check(app.status_filter_menu, "status filter")
app.status_filter_menu._close()
app.update()

# theme toggle both ways
for _ in range(2):
    app._toggle_theme()
    app.update()
    time.sleep(0.2)
    app.update()
print("theme toggle: OK, theme =", m.CURRENT_THEME)

# after rebuild the click wiring must still be correct
assert app.days_filter_menu._command == app.days_filter_menu._open_menu
open_check(app.days_filter_menu, "days filter after theme rebuild")
app.days_filter_menu._close()
app.update()
print("click wiring survives theme rebuild: OK")

# category nav
for name, b in list(app.category_buttons.items()):
    b.invoke()
    app.update()
    time.sleep(0.15)
    app.update()
print("category buttons: OK ->", app.current_category if hasattr(app, "current_category") else "")

app.category_buttons["GBS-market"].invoke()
app.update()
time.sleep(0.2)
app.update()

values = app.sheet_seg.cget("values")
assert values, "sheet segmented control must have values"
app.sheet_seg.set(values[1])
app.update()
print("sheet segmented: OK ->", app.sheet_seg.get())

# date picker long lists (scrollable day/month/year)
dlg = m.LicenseCardDialog(app, app.db, license_data=None,
                          sheet_name="GBS", category="GBS-market")
dlg.report_callback_exception = hook
dlg.update()
entry = dlg.entries.get('id_expiry')
assert entry is not None
dlg._pick_date(entry)
dlg.update()
time.sleep(0.2)
dlg.update()

picker = None


def find_picker(widget):
    global picker
    for child in widget.winfo_children():
        if isinstance(child, tk.Toplevel):
            picker = child
        find_picker(child)


find_picker(dlg)
picker = dlg._picker if hasattr(dlg, "_picker") else picker
assert picker is not None, "date picker not found"
picker.report_callback_exception = hook


def collect_dropdowns(widget, out):
    for child in widget.winfo_children():
        if isinstance(child, m.DropdownOption):
            out.append(child)
        collect_dropdowns(child, out)


dropdowns = []
collect_dropdowns(picker, dropdowns)
assert len(dropdowns) >= 3, f"expected 3 date dropdowns, got {len(dropdowns)}"
for widget in dropdowns:
    open_check(widget, f"picker list ({len(widget._values)} items)")
    widget._choose(widget._values[0])
    app.update()
print("date picker dropdowns: OK")

try:
    picker.destroy()
except Exception:
    pass
dlg.destroy()
app.update()
app.destroy()

assert not errors, "Tk callback errors:\n" + "\n".join(errors)
print("ALL DROPDOWN / TOGGLE TESTS PASSED")
