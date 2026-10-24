# -*- coding: utf-8 -*-
"""Categories: card selector, moving records, the new «Отвалы» category."""
import os
import sys
import tempfile
import shutil

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import main as m
from database import Database

assert m.CATEGORIES == ["GBS-market", "Shopuchet", "МойСклад", "Отвалы"], m.CATEGORIES
assert list(m.IMPORT_CATEGORIES) == m.CATEGORIES
print("CATEGORIES CONSTANT OK")

tmpdir = tempfile.mkdtemp(prefix="lt_category_")
app = None
try:
    test_db = Database(os.path.join(tmpdir, "test.db"))

    first = test_db.create_license({
        "full_name": "Первый Клиент", "category": "GBS-market",
        "sheet_name": "GBS", "gbs_id": "1:A",
    })
    second = test_db.create_license({
        "full_name": "Второй Клиент", "category": "GBS-market",
        "sheet_name": "GBS",
    })
    otval = test_db.create_license({
        "full_name": "Отвальный Клиент", "category": "GBS-market",
        "sheet_name": "отвалы",
    })

    # ---------- bulk move ----------
    moved = test_db.move_licenses([first, second], "МойСклад")
    assert moved == 2, moved
    assert len(test_db.get_all_licenses(category="МойСклад")) == 2
    assert len(test_db.get_all_licenses(category="GBS-market")) == 1
    assert test_db.get_license(first)["category"] == "МойСклад"
    print("BULK MOVE OK")

    # ---------- move a whole sheet ----------
    moved = test_db.move_sheet_records("отвалы", "GBS-market", "Отвалы")
    assert moved == 1, moved
    assert test_db.get_license(otval)["category"] == "Отвалы"
    assert len(test_db.get_all_licenses(category="Отвалы")) == 1
    print("SHEET MOVE OK")

    # ---------- app nav ----------
    app = m.LicenseTrackerApp()
    app.geometry("1200x760+60+40")
    app.deiconify()
    app.update()

    assert set(app.category_buttons.keys()) == set(m.CATEGORIES), app.category_buttons.keys()
    assert "Отвалы" in app.category_buttons
    app._on_category_change("Отвалы")
    app.update()
    print("NAV HAS OTVALY OK, rows:", len(app.tree.get_children()))

    # production DB: legacy "отвалы" sheet was moved into the new category once
    production_moved = app.db.get_all_licenses(category="Отвалы")
    assert production_moved, "«Отвалы» should contain the migrated records"
    assert not app.db.get_all_licenses(category="GBS-market", sheet_name="отвалы"), \
        "legacy sheet must not stay in GBS-market"
    print("PRODUCTION MIGRATION OK ->", len(production_moved), "records")

    # ---------- card category selector ----------
    lic = test_db.get_license(first)  # currently МойСклад
    card = m.LicenseCardDialog(app, test_db, lic, category="МойСклад")
    card.update()
    assert card.category_var.get() == "МойСклад", card.category_var.get()
    assert list(card.category_menu._values) == m.CATEGORIES
    print("CARD CATEGORY SELECTOR OK")

    card.category_menu._choose("Отвалы")
    assert card.category_var.get() == "Отвалы"
    data = card._collect_data()
    assert data["category"] == "Отвалы", data["category"]
    test_db.update_license(first, data)
    assert test_db.get_license(first)["category"] == "Отвалы"
    card.destroy()
    print("CARD MOVE ON SAVE OK")

    # ---------- move dialog ----------
    results = []
    dlg = m.MoveCategoryDialog(app, 2, "GBS-market", lambda c: results.append(c))
    dlg.update()
    assert dlg.category_var.get() == "GBS-market"
    assert dlg.move_btn.cget("state") == "disabled"
    dlg.category_var.set("Shopuchet")
    dlg._on_selected("Shopuchet")
    assert dlg.move_btn.cget("state") == "normal"
    assert "Shopuchet" in dlg.hint_label.cget("text")
    dlg._confirm()
    assert results == ["Shopuchet"], results
    print("MOVE DIALOG OK")

    # ---------- app-level migration method ----------
    mig_db = Database(os.path.join(tmpdir, "mig.db"))
    mig_id = mig_db.create_license({
        "full_name": "Миграция", "category": "GBS-market", "sheet_name": "отвалы",
    })
    old_db = app.db
    app.db = mig_db
    app._migrate_otvaly_category()
    assert mig_db.get_license(mig_id)["category"] == "Отвалы"
    # idempotent: running again changes nothing and does not fail
    app._migrate_otvaly_category()
    assert mig_db.get_license(mig_id)["category"] == "Отвалы"
    app.db = old_db
    print("APP MIGRATION METHOD OK")

finally:
    try:
        if app is not None:
            app.destroy()
    except Exception:
        pass
    shutil.rmtree(tmpdir, ignore_errors=True)

print("ALL CATEGORY TESTS PASSED")
