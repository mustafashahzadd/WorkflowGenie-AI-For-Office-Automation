"""
Check database tables and show detailed structure
"""

import sqlite3

# Connect to database
conn = sqlite3.connect('workflowgenie.db')
cursor = conn.cursor()

print("\n" + "="*80)
print("📊 WORKFLOWGENIE DATABASE STRUCTURE")
print("="*80)

# Get all table names
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()

if not tables:
    print("\n❌ No tables found! Run: python create_tables.py")
    conn.close()
    exit()

print(f"\n✅ Found {len(tables)} tables")

# Show structure of each table
for table in tables:
    table_name = table[0]
    
    print(f"\n{'='*80}")
    print(f"📋 TABLE: {table_name}")
    print(f"{'='*80}")
    
    # Get column info
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = cursor.fetchall()
    
    print(f"\n{'Column Name':<20} {'Type':<15} {'Nullable':<10} {'Default':<15} {'Primary Key'}")
    print("-" * 80)
    
    for col in columns:
        col_id, name, col_type, not_null, default_val, pk = col
        nullable = "NO" if not_null else "YES"
        default = str(default_val) if default_val else "-"
        is_pk = "✓" if pk else ""
        
        print(f"{name:<20} {col_type:<15} {nullable:<10} {default:<15} {is_pk}")
    
    # Show foreign keys
    cursor.execute(f"PRAGMA foreign_key_list({table_name})")
    fks = cursor.fetchall()
    
    if fks:
        print(f"\n🔗 Foreign Keys:")
        for fk in fks:
            print(f"  - {fk[3]} → {fk[2]}.{fk[4]}")
    
    # Show indexes
    cursor.execute(f"PRAGMA index_list({table_name})")
    indexes = cursor.fetchall()
    
    if indexes:
        print(f"\n🔍 Indexes:")
        for idx in indexes:
            unique = "UNIQUE" if idx[2] else "NON-UNIQUE"
            print(f"  - {idx[1]} ({unique})")
    
    # Show row count
    cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
    count = cursor.fetchone()[0]
    print(f"\n📊 Total rows: {count}")
    
    # Show sample data if exists
    if count > 0:
        cursor.execute(f"SELECT * FROM {table_name} LIMIT 3")
        samples = cursor.fetchall()
        print(f"\n📝 Sample data (first 3 rows):")
        for i, row in enumerate(samples, 1):
            print(f"  Row {i}: {row[:3]}..." if len(row) > 3 else f"  Row {i}: {row}")

print("\n" + "="*80)
print("✅ Database structure check complete!")
print("="*80 + "\n")

conn.close()