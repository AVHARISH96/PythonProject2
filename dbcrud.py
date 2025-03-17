import tkinter as tk
from tkinter import messagebox
import pyodbc

def connect_db():
    conn_str = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\avhar\OneDrive\Desktop\Emp.accdb;'

    conn = pyodbc.connect(conn_str)
    return conn

def create_record():
    name = entry_name.get()
    position = entry_position.get()
    salary = entry_salary.get()

    if name and position and salary:
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Employees (Name, Position, Salary) VALUES (?, ?, ?)", (name, position, salary))
        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Record added successfully!")
        clear_entries()
        read_records()
    else:
        messagebox.showwarning("Input Error", "Please fill all fields")

def read_records():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Employees")
    rows = cursor.fetchall()
    conn.close()

    listbox.delete(0, tk.END)
    for row in rows:
        listbox.insert(tk.END, row)

def update_record():
    selected = listbox.curselection()
    print(selected)
    if selected:
        id = listbox.get(selected).split(',')[0].replace('(','')
        name = entry_name.get()
        position = entry_position.get()
        salary = entry_salary.get()

        if name and position and salary:
            conn = connect_db()
            cursor = conn.cursor()
            cursor.execute("UPDATE Employees SET Name=?, Position=?, Salary=? WHERE ID=?", (name, position, salary, id))
            conn.commit()
            conn.close()
            messagebox.showinfo("Success", "Record updated successfully!")
            clear_entries()
            read_records()
        else:
            messagebox.showwarning("Input Error", "Please fill all fields")
    else:
        messagebox.showwarning("Selection Error", "Please select a record to update")

def delete_record():
    selected = listbox.curselection()
    if selected:
        id = listbox.get(selected).split(',')[0].replace('(','')
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM Employees WHERE ID=?", (id,))
        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Record deleted successfully!")
        clear_entries()
        read_records()
    else:
        messagebox.showwarning("Selection Error", "Please select a record to delete")

def clear_entries():
    entry_name.delete(0, tk.END)
    entry_position.delete(0, tk.END)
    entry_salary.delete(0, tk.END)

def on_select(event):
    clear_entries()
    selected = listbox.curselection()
    if selected:
        record = listbox.get(selected)

        name=record.split(",")[1].replace("'","").strip()
        position = record.split(",")[2].replace("'", "").strip()
        salary = record.split(",")[3].replace(")", "").strip()
        entry_name.insert(0, name)
        entry_position.insert(0, position)
        entry_salary.insert(0, salary)

root = tk.Tk()
root.geometry("800x600")
root.title("Employee Management Portal")

tk.Label(root, text="Name").grid(row=0, column=0)
tk.Label(root, text="Position").grid(row=1, column=0)
tk.Label(root, text="Salary").grid(row=2, column=0)

entry_name = tk.Entry(root)
entry_name.grid(row=0, column=1)
entry_position = tk.Entry(root)
entry_position.grid(row=1, column=1)
entry_salary = tk.Entry(root)
entry_salary.grid(row=2, column=1)

tk.Button(root, text="Create", command=create_record).grid(row=3, column=0)
tk.Button(root, text="Update", command=update_record).grid(row=3, column=2)
tk.Button(root, text="Delete", command=delete_record).grid(row=3, column=3)

listbox = tk.Listbox(root, width=50)
listbox.grid(row=4, column=0, columnspan=4)
listbox.bind('<<ListboxSelect>>', on_select)

read_records()

root.mainloop()