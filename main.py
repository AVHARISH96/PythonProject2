import pyodbc

conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\avhar\OneDrive\Desktop\Stud.accdb;'

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

def create_record():
    cursor.execute("INSERT INTO Student (Name,Marks) VALUES (?, ?)", ('Harish', '32'))
    conn.commit()
    print("Record created")

def read_records():
        cursor.execute("SELECT * FROM Student")
        rows = cursor.fetchall()
        for row in rows:
            print(row)

def update_record():
    cursor.execute("UPDATE Student SET Marks = ? WHERE Name = ?", ('95', 'Sam'))
    conn.commit()
    print("Record updated")


def delete_record():
    cursor.execute("DELETE FROM Student WHERE ID = ?", ('3',))
    conn.commit()
    print("Record deleted")

def close_connection():
        cursor.close()
        conn.close()
        print("Connection closed")

read_records()