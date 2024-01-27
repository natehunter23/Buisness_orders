import sys
import sqlite3
import pandas as pd

# Connect to the SQLite database
conn = sqlite3.connect('customer_orders.db')
cursor = conn.cursor()

# Create a table to store customer information and orders
cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        first_name TEXT,
        last_name TEXT,
        order_details TEXT,
        order_price REAL,
        order_status TEXT
    )
''')
conn.commit()

def add_order(first_name, last_name, order_details, order_price, order_status):
    # Add a new order to the database
    cursor.execute('''
        INSERT INTO orders (first_name, last_name, order_details, order_price, order_status)
        VALUES (?, ?, ?, ?, ?)
    ''', (first_name, last_name, order_details, order_price, order_status))
    conn.commit()

def search_orders(column, value):
    # Search for orders based on the specified column and value
    cursor.execute(f'SELECT * FROM orders WHERE {column} = ?', (value,))
    results = cursor.fetchall()
    return results

def update_order(order_id, column, new_value):
    # Update an existing order
    cursor.execute(f'UPDATE orders SET {column} = ? WHERE id = ?', (new_value, order_id))
    conn.commit()

def export_to_excel():
    # Fetch all orders from the database
    cursor.execute('SELECT * FROM orders')
    results = cursor.fetchall()

    # Create a DataFrame from the results
    df = pd.DataFrame(results, columns=['ID', 'First Name', 'Last Name', 'Order Details', 'Order Price', 'Order Status'])

    # Export the DataFrame to an Excel spreadsheet
    df.to_excel('customer_orders.xlsx', index=False)

# Command-line interface
if len(sys.argv) < 2:
    print("Usage: python program.py <action> <arguments>")
    sys.exit(1)

action = sys.argv[1].lower()

if action == 'add':
    if len(sys.argv) != 7:
        print("Usage: python program.py add <first_name> <last_name> <order_details> <order_price> <order_status>")
        sys.exit(1)
    add_order(*sys.argv[2:])
    print("Order added successfully.")

elif action == 'search':
    if len(sys.argv) != 4:
        print("Usage: python program.py search <column> <value>")
        sys.exit(1)
    results = search_orders(sys.argv[2], sys.argv[3])
    print("Search Result:")
    print(results)

elif action == 'update':
    if len(sys.argv) != 5:
        print("Usage: python program.py update <order_id> <column> <new_value>")
        sys.exit(1)
    update_order(*sys.argv[2:])
    print("Order updated successfully.")

elif action == 'export':
    if len(sys.argv) != 2:
        print("Usage: python program.py export")
        sys.exit(1)
    export_to_excel()
    print("Data exported to Excel successfully.")

else:
    print("Invalid action. Available actions: add, search, update, export.")
    sys.exit(1)

# Close the database connection
conn.close()
