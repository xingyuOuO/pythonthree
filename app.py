import re
import requests
import sqlite3
from tkinter import Tk, Label, Entry, Button, StringVar, messagebox, ttk

# 預設資料庫文件與 URL
DB_FILE = "contacts.db"
DEFAULT_URL = "https://landscape.ncut.edu.tw/p/412-1031-6821.php"


def setup_database():
    """建立 SQLite 資料表"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute(''' 
        CREATE TABLE IF NOT EXISTS contacts (
            iid INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            title TEXT NOT NULL,
            email TEXT NOT NULL,
            phone_extension TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()


def save_to_database(name: str, title: str, email: str, phone_extension: str) -> None:
    """將聯絡資訊存入資料庫（避免重複記錄）"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        # 先檢查資料是否存在
        cursor.execute('''
            SELECT 1 FROM contacts WHERE name = ? AND email = ?
        ''', (name, email))
        if cursor.fetchone():
            print(f"資料已存在: {name}, {email}")  # 或者用 log 輸出來
            return  # 若資料已經存在則不插入
        
        # 插入資料
        cursor.execute('''
            INSERT INTO contacts (name, title, email, phone_extension)
            VALUES (?, ?, ?, ?)
        ''', (name, title, email, phone_extension))
        conn.commit()
    except sqlite3.Error:
        pass
    finally:
        conn.close()


def parse_contacts(html: str) -> list[tuple[str, str, str, str]]:
    """使用正規表示式從 HTML 中提取聯絡資訊（姓名、職稱、Email、電話）"""
    # 正規表示式，用來抓取姓名、職稱、Email 和電話
    contact_pattern = re.compile(
        r'<a href=".*?"><img.*?alt="(.*?)".*?</a>.*?職　　稱：</td>\s*<td>(.*?)</td>.*?電子郵件 :.*?mailto:(.*?)".*?聯絡電話 :.*?#(\d{4})</td>', 
        re.DOTALL)
    
    matches = contact_pattern.findall(html)
    
    contacts = []
    for match in matches:
        name, title, email, phone_extension = match
        if name and title and email and phone_extension:
            contacts.append((name, title, email, phone_extension))
    return contacts


def scrape_contacts(url: str) -> list[tuple[str, str, str, str]]:
    """發送 HTTP 請求並提取聯絡資訊"""
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return parse_contacts(response.text)
    except requests.exceptions.RequestException as e:
        messagebox.showerror("網路錯誤", f"無法連線至 {url}：{e}")
        return []


def display_contacts(tree: ttk.Treeview, contacts: list[tuple[str, str, str, str]]) -> None:
    """在 Tkinter 界面中顯示聯絡資訊"""
    tree.delete(*tree.get_children())  # 清空表格
    for contact in contacts:
        tree.insert('', 'end', values=contact)


def main():
    """主程式，建立 Tkinter 界面並執行應用程式"""
    setup_database()

    def on_scrape():
        url = url_var.get()
        contacts = scrape_contacts(url)
        if contacts:
            tree.delete(*tree.get_children())
            for name, title, email, phone_extension in contacts:
                save_to_database(name, title, email, phone_extension)
            display_contacts(tree, contacts)

    root = Tk()
    root.title("聯絡資訊擷取工具")
    root.geometry("800x600")

    # URL 輸入區
    url_var = StringVar(value=DEFAULT_URL)
    Label(root, text="目標 URL：").grid(row=0, column=0, sticky='w', padx=10, pady=10)
    Entry(root, textvariable=url_var, width=50).grid(row=0, column=1, sticky='ew', padx=10, pady=10)
    Button(root, text="抓取", command=on_scrape).grid(row=0, column=2, padx=10, pady=10)

    # 顯示區
    tree = ttk.Treeview(root, columns=("姓名", "職稱", "電子郵件", "分機"), show="headings")
    tree.heading("姓名", text="姓名")
    tree.heading("職稱", text="職稱")
    tree.heading("電子郵件", text="電子郵件")
    tree.heading("分機", text="分機")
    tree.column("姓名", width=150, anchor='w')
    tree.column("職稱", width=150, anchor='w')
    tree.column("電子郵件", width=200, anchor='w')
    tree.column("分機", width=100, anchor='w')
    tree.grid(row=1, column=0, columnspan=3, sticky='nsew', padx=10, pady=10)

    # 調整佈局
    root.grid_columnconfigure(1, weight=1)
    root.grid_rowconfigure(1, weight=1)

    root.mainloop()


if __name__ == "__main__":
    main()
