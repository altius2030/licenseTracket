# -*- coding: utf-8 -*-
"""Excel import: column auto-mapping, record building and the import wizard."""
import sys, os, tempfile, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import main as m
from database import Database

# ---------- 1) auto-mapping ----------
headers = ["Дата установки", "Фамилия и имя", "Название магазина", "Телефон",
           "Почта", "GBS ID", "Срок  ID", "Дата оплаты", "Адрес"]
mapping = m.auto_map_columns(headers)
assert mapping["full_name"] == "Фамилия и имя", mapping
assert mapping["store_name"] == "Название магазина", mapping
assert mapping["phone"] == "Телефон", mapping
assert mapping["email"] == "Почта", mapping
assert mapping["gbs_id"] == "GBS ID", mapping
assert mapping["id_expiry"] == "Срок  ID", mapping
assert mapping["install_date"] == "Дата установки", mapping
assert mapping["payment_date"] == "Дата оплаты", mapping
assert mapping["address"] == "Адрес", mapping
print("AUTO-MAPPING OK")

# exact matching must give 'Оплата' to the text field, not the date field
map2 = m.auto_map_columns(["ФИО", "Оплата"])
assert map2["payment_info"] == "Оплата", map2
assert "payment_date" not in map2, map2
print("AMBIGUOUS HEADERS OK")

# ---------- 2) building records ----------
df = pd.DataFrame([
    {"ФИО": "Джаикбаева Камила", "Почта": "Dzhaikbaeva@bk.ru",
     "Название магазина": "Магазин одежды", "Телефон": 87473465436.0,
     "GBS ID": "6:E61991B5D4BB4173A3301A899A31EF77", "Срок ID": "15.01.2027"},
    {"ФИО": None, "Почта": None, "Название магазина": None, "Телефон": None,
     "GBS ID": None, "Срок ID": None},
])
recs = m.build_records_from_frame(df, {
    "full_name": "ФИО", "email": "Почта", "store_name": "Название магазина",
    "phone": "Телефон", "gbs_id": "GBS ID", "id_expiry": "Срок ID",
})
assert len(recs) == 1, recs
rec = recs[0]
assert rec["phone"] == "87473465436", rec["phone"]
assert rec["id_expiry"] == "2027-01-15", rec["id_expiry"]
assert rec["full_name"] == "Джаикбаева Камила"
print("RECORD BUILDING OK")

# ---------- 3) wizard (real xlsx file) ----------
tmpdir = tempfile.mkdtemp(prefix="lt_import_")
try:
    import openpyxl
    xlsx = os.path.join(tmpdir, "клиенты.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Клиенты"
    ws.append(["ФИО", "Почта", "Название магазина", "Телефон", "GBS ID"])
    ws.append(["Джаикбаева Камила", "Dzhaikbaeva@bk.ru", "Магазин одежды",
               "8747 346 54 36", "6:E61991B5D4BB4173A3301A899A31EF77"])
    ws.append(["Иванов Олег", "ivanov@mail.ru", "Автозапчасти",
               "87013780707", "6:B5FC3DB458B64066A0CE086AF527833F"])
    wb.save(xlsx)

    app = m.LicenseTrackerApp()
    app.geometry("1200x760+80+40")
    app.deiconify()
    app.update()

    # category nav must be centered in the top bar
    place_info = app.category_inner.place_info()
    assert float(place_info.get("relx", 0)) == 0.5, place_info
    print("CENTERED CATEGORY NAV OK")

    infos = []
    m.messagebox.showinfo = lambda *a, **k: infos.append(a)
    m.messagebox.showwarning = lambda *a, **k: infos.append(a)
    m.messagebox.showerror = lambda *a, **k: (_ for _ in ()).throw(AssertionError(a))

    test_db = Database(os.path.join(tmpdir, "test.db"))
    dlg = m.ImportDialog(app, test_db, default_category="Shopuchet")
    dlg.update()

    dlg.file_path = xlsx
    dlg.file_label.configure(text=os.path.basename(xlsx))
    dlg._load_preview()
    dlg.update()

    assert dlg.preview_card.winfo_ismapped(), "preview card must be shown"
    assert dlg.rows_label.cget("text") == "строк в файле: 2", dlg.rows_label.cget("text")
    assert dlg.sheet_var.get() == "Клиенты", dlg.sheet_var.get()
    assert dlg.field_vars["full_name"].get() == "ФИО"
    assert dlg.field_vars["email"].get() == "Почта"
    assert dlg.field_vars["store_name"].get() == "Название магазина"
    assert dlg.field_vars["phone"].get() == "Телефон"
    assert dlg.field_vars["gbs_id"].get() == "GBS ID"
    assert dlg.field_vars["address"].get() == m.NO_COLUMN
    assert dlg.category_var.get() == "Shopuchet"
    assert dlg.import_btn.cget("state") == "normal"
    assert dlg.category_seg.pack_info().get("anchor") == "center"
    print("WIZARD PREVIEW + MAPPING OK")

    # user changes the category, then imports
    dlg.category_var.set("GBS-market")
    dlg._do_import()
    app.update()

    rows = test_db.get_all_licenses(category="GBS-market")
    assert len(rows) == 2, f"expected 2 imported rows, got {len(rows)}"
    imported = {r["full_name"]: r for r in rows}
    kamila = imported["Джаикбаева Камила"]
    assert kamila["email"] == "Dzhaikbaeva@bk.ru"
    assert kamila["store_name"] == "Магазин одежды"
    assert kamila["phone"] == "8747 346 54 36"
    assert kamila["gbs_id"] == "6:E61991B5D4BB4173A3301A899A31EF77"
    assert kamila["sheet_name"] == "Клиенты"
    assert kamila["category"] == "GBS-market"
    assert infos, "import result message must be shown"
    print("IMPORT INTO CHOSEN CATEGORY OK ->", kamila["full_name"])

    # the copied block from the imported card matches the requested format
    block = m.format_card_copy_text(kamila)
    assert block == ("Джаикбаева Камила\nDzhaikbaeva@bk.ru\nМагазин одежды\n"
                     "8747 346 54 36\n6:E61991B5D4BB4173A3301A899A31EF77"), block
    print("IMPORTED CARD COPIES CORRECTLY OK")

    app.destroy()
finally:
    shutil.rmtree(tmpdir, ignore_errors=True)

print("ALL IMPORT TESTS PASSED")
