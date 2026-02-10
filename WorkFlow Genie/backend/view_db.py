import sqlite3

# Connect to database
conn = sqlite3.connect('workflow_genie.db')
cursor = conn.cursor()

# Query 1: View all users
print("\n=== USERS ===")
cursor.execute("SELECT id, username, created_at, last_login FROM users")
for row in cursor.fetchall():
    print(f"ID: {row[0]}, Username: {row[1]}, Created: {row[2]}, Last Login: {row[3]}")

# Query 2: View sessions with user link
print("\n=== SESSIONS (with user) ===")
cursor.execute("""
SELECT 
    users.username,
    sessions.id,
    sessions.created_at,
    sessions.summary
FROM sessions
JOIN users ON sessions.user_id = users.id
ORDER BY sessions.created_at DESC
LIMIT 5
""")
for row in cursor.fetchall():
    print(f"User: {row[0]}, Session: {row[1]}, Created: {row[2]}, Summary: {row[3]}")

# Query 3: View tasks
print("\n=== TASKS (with prompts and responses) ===")
cursor.execute("""
SELECT 
    id,
    prompt,
    status,
    progress,
    response,
    created_at,
    completed_at
FROM tasks
ORDER BY created_at DESC
LIMIT 5
""")
for row in cursor.fetchall():
    print(f"\nTask ID: {row[0]}")
    print(f"Prompt: {row[1][:60]}...")
    print(f"Status: {row[2]}, Progress: {row[3]}%")
    print(f"Response: {row[4][:60] if row[4] else 'None'}...")
    print(f"Created: {row[5]}, Completed: {row[6]}")

# Query 4: View user -> session -> task link
print("\n=== USER -> SESSION -> TASK LINK ===")
cursor.execute("""
SELECT 
    users.id as user_id,
    users.username,
    sessions.id as session_id,
    tasks.prompt,
    tasks.response,
    sessions.summary,
    tasks.created_at
FROM users
JOIN sessions ON users.id = sessions.user_id
LEFT JOIN tasks ON sessions.id = tasks.session_id
WHERE tasks.id IS NOT NULL
ORDER BY tasks.created_at DESC
LIMIT 5
""")
for row in cursor.fetchall():
    print(f"\nUser {row[0]} ({row[1]}) -> Session {row[2]}")
    print(f"Prompt: {row[3][:50]}...")
    print(f"Response: {row[4][:50] if row[4] else 'None'}...")
    print(f"Summary: {row[5][:50] if row[5] else 'None'}...")

conn.close()
print("\n✅ Database query complete!")