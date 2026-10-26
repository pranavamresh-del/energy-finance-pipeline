import csv
import sqlite3

db_name = "energy_financials.db"

try:
    conn = sqlite3.connect(db_name)
    cursor = conn.cursor()

    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [r[0] for r in cursor.fetchall() if not r[0].startswith("sqlite_")]

    print("\n========================================")
    print("FOUND TABLES IN DATABASE:")
    print(tables)
    print("========================================\n")

    if not tables:
        print("⚠️ Warning: No tables found in", db_name)
    else:
        for table in tables:
            cursor.execute(f'SELECT * FROM "{table}"')
            col_names = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            filename = f"{table}.csv"
            with open(filename, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(col_names)
                writer.writerows(rows)
            print(f"✅ Created {filename} ({len(rows)} rows)")

    conn.close()
    print("\nDone! Check your folder for the new .csv files.")

except Exception as e:
    print("\n❌ ERROR OCCURRED:", e)