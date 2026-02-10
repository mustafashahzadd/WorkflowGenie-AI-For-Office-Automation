import sqlite3

# Connect to database
conn = sqlite3.connect('workflow_genie.db')
cursor = conn.cursor()

# Get all table names
print("\n=== TABLES IN DATABASE ===")
cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
tables = cursor.fetchall()
for table in tables:
    print(f"- {table[0]}")

# Show structure of each table
for table in tables:
    table_name = table[0]
    print(f"\n=== Structure of {table_name} ===")
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = cursor.fetchall()
    for col in columns:
        print(f"  {col[1]} ({col[2]})")
    
    # Show count
    cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
    count = cursor.fetchone()[0]
    print(f"  Total rows: {count}")

conn.close()