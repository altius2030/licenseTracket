import sys
sys.path.insert(0, r'C:\Users\user\Desktop\Лицензии\LicenseTracker')
from database import Database
from datetime import date, datetime as dt

db = Database()
lics = db.get_all_licenses(category='GBS-market')

today = date.today()
counts = {'0-3': 0, '5': 0, '7': 0, '10': 0, '30': 0, 'expired': 0, 'none': 0}
for l in lics:
    days = None
    for f in ['next_payment', 'id_expiry']:
        v = l.get(f)
        if v:
            try:
                exp = dt.strptime(str(v)[:10], '%Y-%m-%d').date()
                days = (exp - today).days
                break
            except:
                pass
    if days is None:
        counts['none'] += 1
    elif days < 0:
        counts['expired'] += 1
    elif days <= 3:
        counts['0-3'] += 1
    elif days <= 5:
        counts['5'] += 1
    elif days <= 7:
        counts['7'] += 1
    elif days <= 10:
        counts['10'] += 1
    elif days <= 30:
        counts['30'] += 1

print('GBS-market breakdown:')
for k, v in counts.items():
    print(f'  {k}: {v}')

# Test Shopuchet
lics2 = db.get_all_licenses(category='Shopuchet')
counts2 = {'0-3': 0, '5': 0, '7': 0, '10': 0, '30': 0, 'expired': 0, 'none': 0}
for l in lics2:
    days = None
    for f in ['next_payment', 'id_expiry']:
        v = l.get(f)
        if v:
            try:
                exp = dt.strptime(str(v)[:10], '%Y-%m-%d').date()
                days = (exp - today).days
                break
            except:
                pass
    if days is None:
        counts2['none'] += 1
    elif days < 0:
        counts2['expired'] += 1
    elif days <= 3:
        counts2['0-3'] += 1
    elif days <= 5:
        counts2['5'] += 1
    elif days <= 7:
        counts2['7'] += 1
    elif days <= 10:
        counts2['10'] += 1
    elif days <= 30:
        counts2['30'] += 1

print('Shopuchet breakdown:')
for k, v in counts2.items():
    print(f'  {k}: {v}')
print('OK')
