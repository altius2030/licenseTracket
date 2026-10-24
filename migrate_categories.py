"""
Migration script to assign existing records to categories.
Maps Excel sheets to categories:
- GBS-market: GBS, АСТАНА, Алматы, отвалы
- Shopuchet: ШопУчет, шоп учет
"""
import sqlite3
import sys
sys.path.insert(0, r'C:\Users\user\Desktop\Лицензии\LicenseTracker')

def migrate():
    conn = sqlite3.connect(r'C:\Users\user\Desktop\Лицензии\LicenseTracker\licenses.db')
    cursor = conn.cursor()
    
    # Check if records already have category
    cursor.execute("PRAGMA table_info(licenses)")
    columns = [col[1] for col in cursor.fetchall()]
    print('Current columns:', columns)
    
    if 'category' not in columns:
        print('Adding category column...')
        cursor.execute("ALTER TABLE licenses ADD COLUMN category TEXT NOT NULL DEFAULT 'GBS-market'")
        conn.commit()
    
    # Map sheets to categories
    sheet_to_category = {
        'GBS': 'GBS-market',
        'АСТАНА': 'GBS-market',
        'Алматы': 'GBS-market',
        'отвалы': 'GBS-market',
        'ШопУчет': 'Shopuchet',
        'шоп учет': 'Shopuchet',
    }
    
    total_updated = 0
    for sheet, category in sheet_to_category.items():
        cursor.execute(
            "UPDATE licenses SET category = ? WHERE sheet_name = ?",
            (category, sheet)
        )
        updated = cursor.rowcount
        total_updated += updated
        print(f'  {sheet} -> {category}: {updated} records')
    
    conn.commit()
    
    # Verify
    cursor.execute("SELECT category, COUNT(*) FROM licenses GROUP BY category")
    print('\nFinal distribution:')
    for cat, count in cursor.fetchall():
        print(f'  {cat}: {count}')
    
    conn.close()
    print(f'\nMigration complete. Total records updated: {total_updated}')

if __name__ == "__main__":
    migrate()