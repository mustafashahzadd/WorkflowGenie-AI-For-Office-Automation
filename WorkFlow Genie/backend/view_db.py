import sqlite3

# Connect to database
conn = sqlite3.connect('workflowgenie.db')
cursor = conn.cursor()

# Get all tables in the database (includes user and SQLite system tables)
cursor.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
tables = [row[0] for row in cursor.fetchall()]

if not tables:
    print("No tables found in the database.")
else:
    for table_name in tables:
        print(f"\n=== TABLE: {table_name} ===")

        cursor.execute(f'PRAGMA table_info("{table_name}")')
        columns_info = cursor.fetchall()
        column_names = [col[1] for col in columns_info]

        if column_names:
            print(f"Columns: {', '.join(column_names)}")
        else:
            print("Columns: (none)")

        cursor.execute(f'SELECT * FROM "{table_name}"')
        rows = cursor.fetchall()

        if not rows:
            print("Rows: (empty table)")
            continue

        for i, row in enumerate(rows, start=1):
            print(f"Row {i}: {row}")

conn.close()
print("\nDatabase query complete!")