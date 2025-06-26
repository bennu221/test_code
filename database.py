import sqlite3
import os
import openpyxl

def init_db():
    conn = sqlite3.connect("passwords.db")
    cursor = conn.cursor()
    cursor.execute("""CREATE TABLE IF NOT EXISTS passwords (
                        id INTEGER PRIMARY KEY,
                        website TEXT,
                        username TEXT,
                        password BLOB)""")
    conn.commit()
    conn.close()

def save_password(website, username, password):
    conn = sqlite3.connect("passwords.db")
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM passwords WHERE website = ?", (website,))
    exists = cursor.fetchone()
    if exists:
        conn.close()
        return False  # Already exists
    cursor.execute("INSERT INTO passwords (website, username, password) VALUES (?, ?, ?)",
                   (website, username, password))
    conn.commit()
    conn.close()
    save_to_excel(website, username, password)
    return True

def save_to_excel(website, username, password):
    excel_path = "passwords.xlsx"
    if not os.path.exists(excel_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Website", "Username", "Password"])
        wb.save(excel_path)
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    ws.append([website, username, password])
    wb.save(excel_path)

def retrieve_password(website):
    conn = sqlite3.connect("passwords.db")
    cursor = conn.cursor()
    cursor.execute("SELECT username, password FROM passwords WHERE website=?", (website,))
    result = cursor.fetchone()
    conn.close()
    return result

def get_all_passwords():
    conn = sqlite3.connect("passwords.db")
    c = conn.cursor()
    c.execute("SELECT website, username, password FROM passwords")
    rows = c.fetchall()
    conn.close()
    return rows

def delete_password(website):
    conn = sqlite3.connect("passwords.db")
    c = conn.cursor()
    c.execute("DELETE FROM passwords WHERE website = ?", (website,))
    conn.commit()
    conn.close()

def update_password(website, new_password):
    conn = sqlite3.connect("passwords.db")
    c = conn.cursor()
    c.execute("UPDATE passwords SET password = ? WHERE website = ?", (new_password, website))
    conn.commit()
    conn.close()
