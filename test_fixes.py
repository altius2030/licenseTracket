import sys
sys.path.insert(0, r'C:\Users\user\Desktop\Лицензии\LicenseTracker')
from database import Database, get_app_dir, parse_excel_date
from datetime import date, timedelta

# Test 1: timedelta import fix (Excel serial dates)
result = parse_excel_date(45000)  # Excel serial date
print('Test 1 - Excel serial date parse:', result)

# Test 2: get_app_dir
print('Test 2 - App dir:', get_app_dir())

# Test 3: DB path
db = Database()
print('Test 3 - DB path:', db.db_path)

# Test 4: Category filtering
gb = db.get_all_licenses(category='GBS-market')
sh = db.get_all_licenses(category='Shopuchet')
my = db.get_all_licenses(category='МойСклад')
print(f'Test 4 - GBS-market: {len(gb)}, Shopuchet: {len(sh)}, МойСклад: {len(my)}')

# Test 5: get_sheets by category
print('Test 5 - GBS sheets:', db.get_sheets('GBS-market'))
print('Test 5 - Shop sheets:', db.get_sheets('Shopuchet'))

# Test 6: update_license with status recalc
if gb:
    lic = gb[0]
    old_status = lic['status']
    db.update_license(lic['id'], {'comment': 'test update', 'category': lic['category'], 'sheet_name': lic['sheet_name']})
    updated = db.get_license(lic['id'])
    print(f'Test 6 - Status preserved: {old_status} -> {updated["status"]}')

# Test 7: date formats
print('Test 7 - Date parse formats:')
print('  DD.MM.YYYY:', parse_excel_date('25.12.2026'))
print('  ISO:', parse_excel_date('2026-12-25'))
print('  datetime obj:', parse_excel_date(date(2026, 12, 25)))
print('  None:', parse_excel_date(None))

# Test 8: empty field clearing
print('Test 8 - Clear field test')
lic2 = gb[1]
db.update_license(lic2['id'], {'comment': None, 'category': lic2['category'], 'sheet_name': lic2['sheet_name']})
check = db.get_license(lic2['id'])
print('  comment cleared:', check['comment'] is None)

# Test 9: create + delete
new_id = db.create_license({
    'sheet_name': 'GBS', 'category': 'GBS-market',
    'full_name': 'Test User', 'store_name': 'Test Store',
    'next_payment': '2027-01-01'
})
print('Test 9 - Created ID:', new_id)
print('  Status:', db.get_license(new_id)['status'])
db.delete_license(new_id)
print('  Deleted:', db.get_license(new_id) is None)

# Test 10: export
import tempfile, os
tmp = os.path.join(tempfile.gettempdir(), 'test_export.xlsx')
db.export_to_excel(tmp, category='GBS-market')
print('Test 10 - Export OK:', os.path.exists(tmp))
os.remove(tmp)

print('\n=== ALL TESTS PASSED ===')
