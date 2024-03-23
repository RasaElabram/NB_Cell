import os
import sqlite3
import csv

def create_database(db_file):
    # Connect to SQLite database (creates it if not existing)
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Get list of tables in the database (excluding system tables)
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
    tables = cursor.fetchall()

    # Drop existing tables
    for table in tables:
        cursor.execute(f"DROP TABLE IF EXISTS {table[0]}")

    # Get list of files in csv_files folder
    csv_files = os.listdir(os.path.join(os.path.dirname(__file__), "..", "csv_files"))

    # Iterate through files to find CSV file
    for file in csv_files:
        if file.endswith(".csv"):
            csv_file = os.path.join(os.path.dirname(__file__), "..", "csv_files", file)
            # Read CSV file to get headers and data
            with open(csv_file, 'r', newline='', encoding='utf-8-sig') as f:  # Specify encoding as 'utf-8-sig' to handle BOM (Byte Order Mark)
                reader = csv.DictReader(f)
                headers = reader.fieldnames

                # Generate column names and types dynamically
                column_definitions = [f"{header.replace(' ', '_').lower()} TEXT" for header in headers]

                # Create table with dynamically generated columns
                cursor.execute(f'''CREATE TABLE IF NOT EXISTS my_table (
                                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                                    {", ".join(column_definitions)}
                                    );''')

                # Insert data into the table
                for row in reader:
                    values = [row[header] for header in headers]
                    cursor.execute("INSERT INTO my_table ({}) VALUES ({});".format(", ".join(headers), ", ".join(['?']*len(headers))), values)

    # Commit changes and close connection
    conn.commit()
    conn.close()

if __name__ == "__main__":
    # Set path for database file
    db_file = os.path.join(os.path.dirname(__file__), "cell_info.db")  # Specify the path to your SQLite database file

    # Create database
    create_database(db_file)
