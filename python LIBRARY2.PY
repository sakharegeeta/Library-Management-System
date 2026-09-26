import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os

# ---------------- Main Window ----------------
root = tk.Tk()
root.title("Library Management System")
root.geometry("800x650")
root.config(bg="lightgray")

# ---------------- Excel File ----------------
file = os.path.join(
    os.path.expanduser("~"),
    "Desktop",
    "LibraryData.xlsx"
)

# ---------------- Variables ----------------
book_id_var = tk.StringVar()
book_name_var = tk.StringVar()
author_var = tk.StringVar()
student_var = tk.StringVar()
search_var = tk.StringVar()

# ---------------- Book Data ----------------
books = {}


# ---------------- Create Excel ----------------
def create_excel():
    if not os.path.exists(file):
        wb = Workbook()
        ws = wb.active

        ws.append([
            "Book ID",
            "Book Name",
            "Author",
            "Status",
            "Student"
        ])

        wb.save(file)


# ---------------- Load Data ----------------
def load_data():
    try:
        wb = load_workbook(file)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):

            if row[0] is not None:
                book_id = str(row[0])

                books[book_id] = {
                    "name": str(row[1] or ""),
                    "author": str(row[2] or ""),
                    "status": str(row[3] or "Available"),
                    "student": str(row[4] or "")
                }

    except:
        create_excel()


# ---------------- Save Data ----------------
def save_data():

    wb = Workbook()
    ws = wb.active

    ws.append([
        "Book ID",
        "Book Name",
        "Author",
        "Status",
        "Student"
    ])

    for book_id, book in books.items():

        ws.append([
            book_id,
            book["name"],
            book["author"],
            book["status"],
            book["student"]
        ])

    wb.save(file)


# ---------------- Add Book ----------------
def add_book():

    book_id = book_id_var.get()
    book_name = book_name_var.get()
    author = author_var.get()

    if book_id == "" or book_name == "" or author == "":
        messagebox.showerror(
            "Error",
            "Please enter all book details"
        )
        return

    if book_id in books:
        messagebox.showerror(
            "Error",
            "Book ID already exists"
        )
        return

    books[book_id] = {
        "name": book_name,
        "author": author,
        "status": "Available",
        "student": ""
    }

    save_data()

    messagebox.showinfo(
        "Success",
        "Book added successfully"
    )

    display_books()
    clear_data()


# ---------------- Issue Book ----------------
def issue_book():

    book_id = book_id_var.get()
    student = student_var.get()

    if book_id == "" or student == "":
        messagebox.showerror(
            "Error",
            "Enter Book ID and Student Name"
        )
        return

    if book_id not in books:
        messagebox.showerror(
            "Error",
            "Book not found"
        )
        return

    if books[book_id]["status"] == "Issued":
        messagebox.showerror(
            "Error",
            "Book is already issued"
        )
        return

    books[book_id]["status"] = "Issued"
    books[book_id]["student"] = student

    save_data()

    messagebox.showinfo(
        "Success",
        "Book issued successfully"
    )

    display_books()


# ---------------- Return Book ----------------
def return_book():

    book_id = book_id_var.get()

    if book_id == "":
        messagebox.showerror(
            "Error",
            "Enter Book ID"
        )
        return

    if book_id not in books:
        messagebox.showerror(
            "Error",
            "Book not found"
        )
        return

    if books[book_id]["status"] == "Available":
        messagebox.showerror(
            "Error",
            "Book is already available"
        )
        return

    books[book_id]["status"] = "Available"
    books[book_id]["student"] = ""

    save_data()

    messagebox.showinfo(
        "Success",
        "Book returned successfully"
    )

    display_books()


# ---------------- Search Book ----------------
def search_book():

    search = search_var.get().lower()

    result_box.delete(1.0, tk.END)

    found = False

    for book_id, book in books.items():

        if (
            search in book_id.lower()
            or search in book["name"].lower()
            or search in book["author"].lower()
        ):

            result_box.insert(
                tk.END,
                "Book ID: " + book_id +
                "\nBook Name: " + book["name"] +
                "\nAuthor: " + book["author"] +
                "\nStatus: " + book["status"] +
                "\nStudent: " + book["student"] +
                "\n" + "-" * 40 + "\n"
            )

            found = True

    if not found:

        result_box.insert(
            tk.END,
            "Book not found"
        )


# ---------------- Display Books ----------------
def display_books():

    result_box.delete(1.0, tk.END)

    if len(books) == 0:

        result_box.insert(
            tk.END,
            "No books available"
        )

        return

    for book_id, book in books.items():

        result_box.insert(
            tk.END,
            "Book ID: " + book_id +
            "\nBook Name: " + book["name"] +
            "\nAuthor: " + book["author"] +
            "\nStatus: " + book["status"] +
            "\nStudent: " + book["student"] +
            "\n" + "-" * 40 + "\n"
        )


# ---------------- Clear Data ----------------
def clear_data():

    book_id_var.set("")
    book_name_var.set("")
    author_var.set("")
    student_var.set("")
    search_var.set("")


# ---------------- Heading ----------------
title = tk.Label(
    root,
    text="LIBRARY MANAGEMENT SYSTEM",
    font=("Arial", 24, "bold"),
    bg="darkblue",
    fg="white",
    pady=15
)

title.pack(fill="x")


# ---------------- Input Frame ----------------
frame = tk.Frame(
    root,
    bg="lightgray"
)

frame.pack(pady=20)


tk.Label(
    frame,
    text="Book ID:",
    font=("Arial", 12),
    bg="lightgray"
).grid(
    row=0,
    column=0,
    padx=10,
    pady=8
)

tk.Entry(
    frame,
    textvariable=book_id_var,
    width=30,
    font=("Arial", 12)
).grid(
    row=0,
    column=1,
    padx=10,
    pady=8
)


tk.Label(
    frame,
    text="Book Name:",
    font=("Arial", 12),
    bg="lightgray"
).grid(
    row=1,
    column=0,
    padx=10,
    pady=8
)

tk.Entry(
    frame,
    textvariable=book_name_var,
    width=30,
    font=("Arial", 12)
).grid(
    row=1,
    column=1,
    padx=10,
    pady=8
)


tk.Label(
    frame,
    text="Author Name:",
    font=("Arial", 12),
    bg="lightgray"
).grid(
    row=2,
    column=0,
    padx=10,
    pady=8
)

tk.Entry(
    frame,
    textvariable=author_var,
    width=30,
    font=("Arial", 12)
).grid(
    row=2,
    column=1,
    padx=10,
    pady=8
)


tk.Label(
    frame,
    text="Student Name:",
    font=("Arial", 12),
    bg="lightgray"
).grid(
    row=3,
    column=0,
    padx=10,
    pady=8
)

tk.Entry(
    frame,
    textvariable=student_var,
    width=30,
    font=("Arial", 12)
).grid(
    row=3,
    column=1,
    padx=10,
    pady=8
)


# ---------------- Buttons ----------------
button_frame = tk.Frame(
    root,
    bg="lightgray"
)

button_frame.pack(pady=10)


tk.Button(
    button_frame,
    text="Add Book",
    command=add_book,
    width=15,
    font=("Arial", 11, "bold")
).grid(
    row=0,
    column=0,
    padx=5
)


tk.Button(
    button_frame,
    text="Issue Book",
    command=issue_book,
    width=15,
    font=("Arial", 11, "bold")
).grid(
    row=0,
    column=1,
    padx=5
)


tk.Button(
    button_frame,
    text="Return Book",
    command=return_book,
    width=15,
    font=("Arial", 11, "bold")
).grid(
    row=0,
    column=2,
    padx=5
)


tk.Button(
    button_frame,
    text="Display Books",
    command=display_books,
    width=15,
    font=("Arial", 11, "bold")
).grid(
    row=0,
    column=3,
    padx=5
)


tk.Button(
    button_frame,
    text="Clear",
    command=clear_data,
    width=15,
    font=("Arial", 11, "bold")
).grid(
    row=1,
    column=1,
    padx=5,
    pady=8
)


# ---------------- Search ----------------
search_frame = tk.Frame(
    root,
    bg="lightgray"
)

search_frame.pack(pady=10)


tk.Label(
    search_frame,
    text="Search Book:",
    font=("Arial", 12, "bold"),
    bg="lightgray"
).grid(
    row=0,
    column=0,
    padx=10
)


tk.Entry(
    search_frame,
    textvariable=search_var,
    width=25,
    font=("Arial", 12)
).grid(
    row=0,
    column=1,
    padx=10
)


tk.Button(
    search_frame,
    text="Search",
    command=search_book,
    width=12,
    font=("Arial", 11, "bold")
).grid(
    row=0,
    column=2,
    padx=10
)


# ---------------- Display Area ----------------
result_box = tk.Text(
    root,
    width=75,
    height=12,
    font=("Arial", 11)
)

result_box.pack(pady=10)


# ---------------- Start Program ----------------
create_excel()
load_data()

root.mainloop()
